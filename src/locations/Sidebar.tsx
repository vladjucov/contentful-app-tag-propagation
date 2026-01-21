import {
  Stack,
  Text,
  Button,
} from '@contentful/f36-components';
import type { SidebarAppSDK } from '@contentful/app-sdk';
import { useSDK } from '@contentful/react-apps-toolkit';
import { useMemo, useRef, useState, type MutableRefObject } from 'react';

function getCurrentContentTypeId(sdk: SidebarAppSDK): string {
  return (sdk as any)?.contentType?.sys?.id ?? '';
}

function parseCsv(value?: string): string[] {
  if (!value) return [];
  return value
    .split(/[\n,]/g)
    .map((s) => s.trim())
    .filter(Boolean);
}

function getDefaultLocale(sdk: SidebarAppSDK): string {
  return sdk.locales.default;
}

type TagMatchMode = 'name' | 'id';

type InstallConfig = {
  excludedContentTypes: string[];
  includeAssetsByDefault: boolean;
  preservePublishStateByDefault: boolean;
  tagMatchMode: TagMatchMode;
  maxTraversalDepth: number;
  requestConcurrency: number;
  rootContentTypes: string[];
  locationLinkFieldIds: string[];
};

function getInstallConfig(sdk: SidebarAppSDK): InstallConfig {
  const p = (sdk.parameters.installation ?? {}) as Record<string, any>;

  return {
    excludedContentTypes: parseCsv(p.excludedContentTypesCsv),
    includeAssetsByDefault: Boolean(p.includeAssetsByDefault ?? true),
    preservePublishStateByDefault: Boolean(p.preservePublishStateByDefault ?? true),
    tagMatchMode: (p.tagMatchMode ?? 'name') as TagMatchMode,
    maxTraversalDepth: Number(p.maxTraversalDepth ?? 15),
    requestConcurrency: Number(p.requestConcurrency ?? 4),
    rootContentTypes: parseCsv(p.rootContentTypesCsv ?? 'pageHome,pageRegular,pageEvent,pageBlogArticle'),
    locationLinkFieldIds: parseCsv(p.locationLinkFieldIdsCsv ?? 'location,locations'),
  };
}

// --- Links extraction (Entry/Asset) ---

type LinkEntry = { sys: { type: 'Link'; linkType: 'Entry'; id: string } };

type LinkAsset = { sys: { type: 'Link'; linkType: 'Asset'; id: string } };

type LocationContext = {
  locationIds: string[];
  locationTags: string[]; // tag names (e.g. "Location: St. Louis")
  warning?: string;
};

function isEntryLink(value: any): value is LinkEntry {
  return Boolean(value?.sys?.type === 'Link' && value?.sys?.linkType === 'Entry' && value?.sys?.id);
}

function isAssetLink(value: any): value is LinkAsset {
  return Boolean(value?.sys?.type === 'Link' && value?.sys?.linkType === 'Asset' && value?.sys?.id);
}

function deepExtractLinks(value: any, out: { entryIds: Set<string>; assetIds: Set<string> }) {
  if (!value) return;

  if (isEntryLink(value)) {
    out.entryIds.add(value.sys.id);
    return;
  }

  if (isAssetLink(value)) {
    out.assetIds.add(value.sys.id);
    return;
  }

  if (Array.isArray(value)) {
    for (const v of value) deepExtractLinks(v, out);
    return;
  }

  if (typeof value === 'object') {
    // RichText, JSON fields, etc.
    for (const k of Object.keys(value)) {
      deepExtractLinks((value as any)[k], out);
    }
  }
}

function getLocalizedFieldValue(field: any, locale: string): any {
  if (!field || typeof field !== 'object') return undefined;
  if (field[locale] !== undefined) return field[locale];
  const keys = Object.keys(field);
  if (keys.length) return field[keys[0]];
  return undefined;
}

function getEntryTitle(entry: any, locale: string): string {
  const fields = entry?.fields ?? {};
  const candidates = ['internalName', 'title', 'name', 'headline'];
  for (const key of candidates) {
    const v = getLocalizedFieldValue(fields?.[key], locale);
    if (typeof v === 'string' && v.trim()) return v;
  }
  return entry?.sys?.id ?? '—';
}

function getTagIdsFromMetadata(entity: any): string[] {
  const tags = entity?.metadata?.tags;
  if (!Array.isArray(tags)) return [];
  return tags
    .map((t: any) => t?.sys?.id)
    .filter((id: any) => typeof id === 'string' && id.length);
}

async function findRootPagesForLocation(
  sdk: SidebarAppSDK,
  locationId: string,
  config: InstallConfig
): Promise<{ rootEntryIds: string[]; perContentType: Record<string, number> }> {
  const rootEntryIds = new Set<string>();
  const perContentType: Record<string, number> = {};

  const limit = 100;

  for (const ct of config.rootContentTypes) {
    const ctRootIds = new Set<string>();

    for (const fieldId of config.locationLinkFieldIds) {
      let skip = 0;

      while (true) {
        try {
          const res = await sdk.cma.entry.getMany({
            spaceId: sdk.ids.space,
            environmentId: sdk.ids.environment,
            query: {
              content_type: ct,
              [`fields.${fieldId}.sys.id`]: locationId,
              limit,
              skip,
            } as any,
          });

          const items: any[] = (res as any)?.items ?? [];
          if (!items.length) break;

          for (const item of items) {
            const id = item?.sys?.id;
            if (id) {
              rootEntryIds.add(id);
              ctRootIds.add(id);
            }
          }

          if (items.length < limit) break;
          skip += limit;
        } catch (e: any) {
          const msg = String(e?.message ?? '');
          // Contentful returns 422 here; we just skip missing fields.
          if (msg.includes('No field with id')) break;
          throw e;
        }
      }
    }

    perContentType[ct] = ctRootIds.size;
  }

  return { rootEntryIds: Array.from(rootEntryIds), perContentType };
}

async function resolveLocationTagsByIds(
  sdk: SidebarAppSDK,
  locationIds: string[],
  locale: string
): Promise<{ tags: string[]; notFound: string[] }> {
  const tags: string[] = [];
  const notFound: string[] = [];

  for (const id of locationIds) {
    try {
      const entry = await sdk.cma.entry.get({
        spaceId: sdk.ids.space,
        environmentId: sdk.ids.environment,
        entryId: id,
      });

      const raw = (entry as any)?.fields?.tag;
      const value = getLocalizedFieldValue(raw, locale);

      if (typeof value === 'string' && value.trim()) {
        tags.push(value);
      } else {
        notFound.push(id);
      }
    } catch {
      notFound.push(id);
    }
  }

  return { tags: Array.from(new Set(tags)), notFound };
}

// Tags (Contentful Tag entity) live on SPACE level (not environment).
async function getAllTagsByName(sdk: SidebarAppSDK): Promise<Map<string, string>> {
  const byName = new Map<string, string>();
  const limit = 100;
  let skip = 0;

  while (true) {
    const res = await (sdk.cma as any).tag.getMany({
      spaceId: sdk.ids.space,
      query: { limit, skip },
    });

    const items: any[] = res?.items ?? [];
    for (const t of items) {
      const id = t?.sys?.id;
      const name = t?.name;
      if (typeof id === 'string' && typeof name === 'string') {
        byName.set(name, id);
      }
    }

    if (items.length < limit) break;
    skip += limit;
  }

  return byName;
}

async function resolveTargetTagIds(
  sdk: SidebarAppSDK,
  tagMatchMode: TagMatchMode,
  tagNamesOrIds: string[],
  tagNameCache: MutableRefObject<Map<string, string> | null>
): Promise<{ tagIds: string[]; missing: string[] }> {
  const uniq = Array.from(new Set(tagNamesOrIds.filter(Boolean)));

  if (tagMatchMode === 'id') {
    return { tagIds: uniq, missing: [] };
  }

  if (!tagNameCache.current) {
    tagNameCache.current = await getAllTagsByName(sdk);
  }

  const missing: string[] = [];
  const tagIds: string[] = [];

  for (const name of uniq) {
    const id = tagNameCache.current.get(name);
    if (id) tagIds.push(id);
    else missing.push(name);
  }

  return { tagIds, missing };
}

// --- Scan result types ---

type EntryScanRow = {
  kind: 'Entry';
  id: string;
  contentType: string;
  title: string;
  excluded: boolean;
  existingTagIds: string[];
  missingTagIds: string[];
  excludedReason?: 'location' | 'excludedContentType';
};

type AssetScanRow = {
  kind: 'Asset';
  id: string;
  title: string;
  excluded: boolean;
  existingTagIds: string[];
  missingTagIds: string[];
};

type ScanSummary = {
  roots: string[];
  targetTagIds: string[];
  missingTargetTagsByName: string[];
  entries: EntryScanRow[];
  assets: AssetScanRow[];
};

async function promisePool<T, R>(
  items: T[],
  concurrency: number,
  worker: (item: T) => Promise<R>
): Promise<R[]> {
  const results: R[] = [];
  let idx = 0;

  const runners = new Array(Math.max(1, concurrency)).fill(0).map(async () => {
    while (idx < items.length) {
      const myIdx = idx++;
      results[myIdx] = await worker(items[myIdx]);
    }
  });

  await Promise.all(runners);
  return results;
}

async function traverseDescendants(
  sdk: SidebarAppSDK,
  locale: string,
  config: InstallConfig,
  rootEntryIds: string[],
  targetTagIds: string[]
): Promise<ScanSummary> {
  const visitedEntries = new Set<string>();
  const visitedAssets = new Set<string>();

  const entryRows = new Map<string, EntryScanRow>();
  const assetRows = new Map<string, AssetScanRow>();

  type QItem = { type: 'Entry'; id: string; depth: number };
  const queue: QItem[] = rootEntryIds.map((id) => ({ type: 'Entry', id, depth: 0 }));

  while (queue.length) {
    const batch = queue.splice(0, Math.max(1, config.requestConcurrency));

    await Promise.all(
      batch.map(async ({ id, depth }) => {
        if (visitedEntries.has(id)) return;
        visitedEntries.add(id);

        const entry = await sdk.cma.entry.get({
          spaceId: sdk.ids.space,
          environmentId: sdk.ids.environment,
          entryId: id,
        });

        const ct = (entry as any)?.sys?.contentType?.sys?.id ?? 'unknown';
        const isExcludedCt = config.excludedContentTypes.includes(ct);
        const isLocationCt = ct === 'location';
        const excluded = isExcludedCt || isLocationCt;
        const excludedReason: EntryScanRow['excludedReason'] = isLocationCt
          ? 'location'
          : isExcludedCt
            ? 'excludedContentType'
            : undefined;
        const title = getEntryTitle(entry, locale);

        const existing = getTagIdsFromMetadata(entry);
        const missing = targetTagIds.filter((t) => !existing.includes(t));

        entryRows.set(id, {
          kind: 'Entry',
          id,
          contentType: ct,
          title,
          excluded,
          excludedReason,
          existingTagIds: existing,
          missingTagIds: excluded ? [] : missing,
        });

        // Stop traversal on excluded nodes or when depth limit reached.
        if (excluded) return;
        if (depth >= config.maxTraversalDepth) return;

        const out = { entryIds: new Set<string>(), assetIds: new Set<string>() };
        const fields = (entry as any)?.fields ?? {};
        for (const fieldId of Object.keys(fields)) {
          const localized = getLocalizedFieldValue(fields[fieldId], locale);
          deepExtractLinks(localized, out);
        }

        for (const nextEntryId of out.entryIds) {
          if (!visitedEntries.has(nextEntryId)) {
            queue.push({ type: 'Entry', id: nextEntryId, depth: depth + 1 });
          }
        }

        if (config.includeAssetsByDefault) {
          for (const assetId of out.assetIds) {
            visitedAssets.add(assetId);
          }
        }
      })
    );
  }

  // Load assets (if enabled)
  if (config.includeAssetsByDefault && visitedAssets.size) {
    const assetIds = Array.from(visitedAssets);
    const assets = await promisePool(assetIds, config.requestConcurrency, async (assetId) => {
      const asset = await sdk.cma.asset.get({
        spaceId: sdk.ids.space,
        environmentId: sdk.ids.environment,
        assetId,
      });

      const title =
        getLocalizedFieldValue((asset as any)?.fields?.title, locale) ??
        getLocalizedFieldValue((asset as any)?.fields?.file, locale)?.fileName ??
        assetId;

      const existing = getTagIdsFromMetadata(asset);
      const missing = targetTagIds.filter((t) => !existing.includes(t));

      const row: AssetScanRow = {
        kind: 'Asset',
        id: assetId,
        title: String(title ?? assetId),
        excluded: false,
        existingTagIds: existing,
        missingTagIds: missing,
      };
      return row;
    });

    for (const r of assets) {
      assetRows.set(r.id, r);
    }
  }

  return {
    roots: rootEntryIds,
    targetTagIds,
    missingTargetTagsByName: [],
    entries: Array.from(entryRows.values()),
    assets: Array.from(assetRows.values()),
  };
}

const Sidebar = () => {
  const sdk = useSDK<SidebarAppSDK>();

  const locale = useMemo(() => getDefaultLocale(sdk), [sdk]);
  const config = useMemo(() => getInstallConfig(sdk), [sdk]);

  const [locationCtx, setLocationCtx] = useState<LocationContext>({ locationIds: [], locationTags: [] });

  // Roots
  const [rootPagesCount, setRootPagesCount] = useState<number>(0);
  const [rootPagesBreakdown, setRootPagesBreakdown] = useState<Record<string, number>>({});
  const [rootPagesByLocation, setRootPagesByLocation] = useState<Record<string, { count: number; breakdown: Record<string, number> }>>({});
  const [rootEntryIds, setRootEntryIds] = useState<string[]>([]);

  // Descendants scan
  const [scanStatus, setScanStatus] = useState<string>('');
  const [scanSummary, setScanSummary] = useState<ScanSummary | null>(null);
  const [isScanning, setIsScanning] = useState(false);
  const tagNameCache = useRef<Map<string, string> | null>(null);

  sdk.window.startAutoResizer();
  const readLocationContext = async (): Promise<LocationContext> => {
    // Location entry itself
    const tagField = (sdk.entry.fields as any)?.tag;
    const directTag = tagField?.getValue?.(locale) ?? tagField?.getValue?.();
    if (typeof directTag === 'string' && directTag.trim()) {
      const ctx: LocationContext = {
        locationIds: [sdk.entry.getSys().id],
        locationTags: [directTag],
      };
      setLocationCtx(ctx);
      return ctx;
    }

    // Any other entry: read linked locations
    const foundLocationIds = new Set<string>();
    for (const fieldId of config.locationLinkFieldIds) {
      const f = (sdk.entry.fields as any)?.[fieldId];
      if (!f?.getValue) continue;
      const v = f.getValue(locale) ?? f.getValue();
      deepExtractLinks(v, { entryIds: foundLocationIds as any, assetIds: new Set<string>() });
      // deepExtractLinks adds entry links into entryIds. We only want Location links; we filter later by reading tags.
    }

    // The above can include non-location entry links; however, `resolveLocationTagsByIds` will only succeed for Location entries with `fields.tag`.
    const locationIds = Array.from(foundLocationIds);

    if (!locationIds.length) {
      const ctx: LocationContext = {
        locationIds: [],
        locationTags: [],
        warning: 'No linked Location found on this entry. The app can only derive target tag(s) from Location.',
      };
      setLocationCtx(ctx);
      return ctx;
    }

    const { tags, notFound } = await resolveLocationTagsByIds(sdk, locationIds, locale);

    const warning = notFound.length
      ? `Some linked Locations do not have a Tag value (or could not be loaded): ${notFound.join(', ')}`
      : undefined;

    const ctx: LocationContext = { locationIds, locationTags: tags, warning };
    setLocationCtx(ctx);
    return ctx;
  };

  const onScan = async () => {
    setIsScanning(true);
    setScanStatus('Starting scan…');

    setRootPagesCount(0);
    setRootPagesBreakdown({});
    setRootPagesByLocation({});
    setRootEntryIds([]);
    setScanSummary(null);

    try {
      const ctx = await readLocationContext();

      if (!ctx.locationTags.length) {
        setScanStatus(ctx.warning ?? 'No Location tag(s) found.');
        return;
      }

      setScanStatus('Resolving target tags…');

      const { tagIds, missing } = await resolveTargetTagIds(
          sdk,
          config.tagMatchMode,
          ctx.locationTags,
          tagNameCache
      );

      if (!tagIds.length) {
        setScanStatus(
            missing.length
                ? `No matching Contentful Tag found for: ${missing.join(', ')}`
                : 'No target tag IDs resolved.'
        );
        return;
      }

      const currentCt = getCurrentContentTypeId(sdk);
      const currentEntryId = sdk.entry.getSys().id;
      const isLocationEntry = currentCt === 'location' && Boolean((sdk.entry.fields as any)?.tag);

      // Page mode
      if (!isLocationEntry) {
        const breakdown: Record<string, number> = { [currentCt || 'currentEntry']: 1 };
        const roots = [currentEntryId];

        setRootPagesByLocation({});
        setRootPagesCount(1);
        setRootPagesBreakdown(breakdown);
        setRootEntryIds(roots);

        setScanStatus('Scanning descendants…');

        const summary = await traverseDescendants(sdk, locale, config, roots, tagIds);
        summary.missingTargetTagsByName = missing;

        setScanSummary(summary);
        setScanStatus('Scan complete.');

        await openDetailsDialog(summary, missing, breakdown, ctx.locationTags, tagIds);
        return;
      }

      // Location mode
      setScanStatus('Finding root pages…');

      const unionRootIds = new Set<string>();
      const unionBreakdown: Record<string, number> = {};
      const byLocation: Record<string, { count: number; breakdown: Record<string, number> }> = {};

      for (const locationId of ctx.locationIds) {
        const { rootEntryIds, perContentType } = await findRootPagesForLocation(sdk, locationId, config);
        byLocation[locationId] = { count: rootEntryIds.length, breakdown: perContentType };

        for (const id of rootEntryIds) unionRootIds.add(id);
        for (const [ct, n] of Object.entries(perContentType)) {
          unionBreakdown[ct] = (unionBreakdown[ct] ?? 0) + n;
        }
      }

      const roots = Array.from(unionRootIds);

      setRootPagesByLocation(byLocation);
      setRootPagesCount(roots.length);
      setRootPagesBreakdown(unionBreakdown);
      setRootEntryIds(roots);

      setScanStatus('Scanning descendants…');

      const summary = await traverseDescendants(sdk, locale, config, roots, tagIds);
      summary.missingTargetTagsByName = missing;

      setScanSummary(summary);
      setScanStatus('Scan complete.');

      await openDetailsDialog(summary, missing, unionBreakdown, ctx.locationTags, tagIds);
    } catch (e: any) {
      setScanStatus(`Scan failed: ${e?.message ?? String(e)}`);
    } finally {
      setIsScanning(false);
    }
  };


  const openDetailsDialog = async (
      summary: ScanSummary,
      missingTargetTagsByName: string[],
      breakdown: Record<string, number>,
      locationTagNames: string[],
      resolvedTargetTagIds: string[]
  ) => {
    const uniqNames = Array.from(new Set(locationTagNames.filter(Boolean)));
    const resolvedByName = tagNameCache.current;

    const targetTags =
        config.tagMatchMode === 'id'
            ? resolvedTargetTagIds.map((id) => ({ id, name: id }))
            : uniqNames
                .map((name) => ({ name, id: resolvedByName?.get(name) }))
                .filter((t): t is { name: string; id: string } => typeof t.id === 'string' && t.id.length > 0)
                .filter((t) => resolvedTargetTagIds.includes(t.id));

    const contentRows = (() => {
      const seen = new Set<string>();
      return summary.entries
        .slice()
        .sort((a, b) => (a.contentType + a.title + a.id).localeCompare(b.contentType + b.title + b.id))
        .map((e) => ({
          id: e.id,
          title: e.title,
          contentType: e.contentType,
          excluded: e.excluded,
          excludedReason: e.excludedReason ?? null,
          willAdd: !e.excluded && e.missingTagIds.length > 0,
          alreadyOk: !e.excluded && e.missingTagIds.length === 0,
        }))
        .filter((r) => {
          if (seen.has(r.id)) return false;
          seen.add(r.id);
          return true;
        });
    })();

    const mediaRows = (() => {
      const seen = new Set<string>();
      return summary.assets
        .slice()
        .sort((a, b) => (a.title + a.id).localeCompare(b.title + b.id))
        .map((a) => ({
          id: a.id,
          title: a.title,
          excluded: a.excluded,
          willAdd: !a.excluded && a.missingTagIds.length > 0,
          alreadyOk: !a.excluded && a.missingTagIds.length === 0,
        }))
        .filter((r) => {
          if (seen.has(r.id)) return false;
          seen.add(r.id);
          return true;
        });
    })();

    const dialogParameters = {
      mode: 'details' as const,
      payload: {
        targetTags,
        targetTagIds: resolvedTargetTagIds,
        missingTargetTagsByName,
        preservePublishState: config.preservePublishStateByDefault,
        contentRows,
        mediaRows,
      },
    };

    await sdk.dialogs.openCurrentApp({
      title: 'Scan details',
      width: 1100,
      minHeight: 760,
      parameters: dialogParameters as any,
    });
  };

  return (
    <Stack flexDirection="column" spacing="spacingM" alignItems="stretch">
      <Stack flexDirection="column" spacing="spacingXs" alignItems="stretch">
        <Button
          variant="primary"
          onClick={onScan}
          isDisabled={isScanning}
          isLoading={isScanning}
          isFullWidth
        >
          Scan for linked content
        </Button>

        <Text fontColor="gray600" fontSize="fontSizeS">
          {scanStatus || 'Ready.'}
        </Text>
      </Stack>

    </Stack>
  );
};

export default Sidebar;
