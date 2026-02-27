import {
  Button,
  Note,
  Notification,
  Paragraph,
  Stack,
  Text,
  Table,
  TableHead,
  TableRow,
  TableCell,
  TableBody,
  Checkbox,
  Badge,
  Flex,
  Tooltip,
  Switch,
  Skeleton,
  TextLink,
} from '@contentful/f36-components';
import type { DialogAppSDK } from '@contentful/app-sdk';
import { useSDK } from '@contentful/react-apps-toolkit';
import { useEffect, useMemo, useState } from 'react';

type ContentRow = {
  id: string;
  title: string;
  contentType: string;
  excluded: boolean;
  excludedReason?: 'location' | 'excludedContentType';
  willAdd: boolean;
  alreadyOk: boolean;
};

type MediaRow = {
  id: string;
  title: string;
  excluded: boolean;
  willAdd: boolean;
  alreadyOk: boolean;
  previewUrl?: string | null;
};

type DialogParams = {
  mode: 'details';
  payload: {
    targetTags: { id: string; name: string }[];
    targetTagIds: string[];
    missingTargetTagsByName: string[];
    preservePublishState: boolean;
    contentRows: ContentRow[];
    mediaRows: MediaRow[];
  };
};

function makeTagLink(id: string) {
  return { sys: { type: 'Link', linkType: 'Tag', id } } as any;
}

async function promisePool<T>(
  items: T[],
  concurrency: number,
  worker: (item: T) => Promise<void>
): Promise<void> {
  let idx = 0;
  const runners = new Array(Math.max(1, concurrency)).fill(0).map(async () => {
    while (idx < items.length) {
      const myIdx = idx++;
      await worker(items[myIdx]);
    }
  });
  await Promise.all(runners);
}

type EntityStatus = 'excluded' | 'published' | 'changed' | 'draft';

function renderStatusBadge(status: EntityStatus, label: string) {
  const variant =
    status === 'published'
      ? 'positive'
      : status === 'changed'
        ? 'primary'
        : status === 'draft'
          ? 'warning'
          : 'featured';
  return <Badge variant={variant as any}>{label}</Badge>;
}

const Dialog = () => {
  const sdk = useSDK<DialogAppSDK>();
  const params = (sdk.parameters.invocation ?? {}) as DialogParams;

  const close = () => sdk.close();

  if ((params as any)?.mode !== 'details') {
    return (
      <Stack flexDirection="column" spacing="spacingM" padding="spacingM">
        <Note variant="warning">No dialog parameters provided.</Note>
        <Button variant="secondary" onClick={close}>
          Close
        </Button>
      </Stack>
    );
  }

  const p = params.payload;

  // Only rows that are actionable (not excluded will add tags)
  const actionableContentRows = useMemo(
    () => p.contentRows.filter((r) => !r.excluded && r.willAdd),
    [p.contentRows]
  );

  const actionableMediaRows = useMemo(
    () => p.mediaRows.filter((r) => !r.excluded && r.willAdd),
    [p.mediaRows]
  );

  const hasAnythingToApply = actionableContentRows.length > 0 || actionableMediaRows.length > 0;

  const [preservePublishState, setPreservePublishState] = useState<boolean>(Boolean(p.preservePublishState));

  const [mediaPreviewById, setMediaPreviewById] = useState<Record<string, string>>({});

  const [statusEntryById, setStatusEntryById] = useState<Record<string, EntityStatus>>({});
  const [statusAssetById, setStatusAssetById] = useState<Record<string, EntityStatus>>({});

  const StatusSkeleton = () => (
    <div style={{ width: 96, height: 24, display: 'inline-block' }}>
      <Skeleton.Container>
        <Skeleton.Image width={96} height={24} radiusX={6} radiusY={6} />
      </Skeleton.Container>
    </div>
  );

  const ImageSkeleton = () => (
    <div style={{ width: 40, height: 40, display: 'inline-block' }}>
      <Skeleton.Container>
        <Skeleton.Image width={40} height={40} radiusX={6} radiusY={6} />
      </Skeleton.Container>
    </div>
  );

  // Load a small set of media previews (and also published status for those assets)
  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      // load at most first 25 previews to keep it fast
      const ids = actionableMediaRows.slice(0, 25).map((r) => r.id).filter(Boolean);
      const next: Record<string, string> = {};
      const nextPublished: Record<string, EntityStatus> = {};

      for (const assetId of ids) {
        if (mediaPreviewById[assetId] && statusAssetById[assetId] !== undefined) continue;
        try {
          const asset = await sdk.cma.asset.get({
            spaceId: sdk.ids.space,
            environmentId: sdk.ids.environment,
            assetId,
          });

          const sys = (asset as any)?.sys;
          const st: EntityStatus =
            !sys?.publishedVersion
              ? 'draft'
              : typeof sys?.version === 'number' && sys.version === sys.publishedVersion + 1
                ? 'published'
                : 'changed';
          nextPublished[assetId] = st;

          const file = (asset as any)?.fields?.file;
          // Try common locales, otherwise first locale.
          const fileVal = (file && (file[sdk.locales.default] ?? file[Object.keys(file)[0]])) || undefined;
          const url = fileVal?.url;
          if (typeof url === 'string' && url.length) {
            next[assetId] = url.startsWith('//') ? `https:${url}?w=40&h=40&fit=thumb` : `${url}?w=40&h=40&fit=thumb`;
          }
        } catch {
          // ignore
        }
      }

      if (!cancelled) {
        if (Object.keys(next).length) {
          setMediaPreviewById((prev) => ({ ...prev, ...next }));
        }
        if (Object.keys(nextPublished).length) {
          setStatusAssetById((prev) => ({ ...prev, ...nextPublished }));
        }
      }
    };

    load();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [actionableMediaRows]);

  // Load published status for more assets (so the Status column can stop showing Skeleton)
  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      const ids = actionableMediaRows.slice(0, 250).map((r) => r.id).filter(Boolean);
      const next: Record<string, EntityStatus> = {};
      const concurrency = 6;

      await promisePool(ids, concurrency, async (assetId) => {
        if (statusAssetById[assetId] !== undefined) return;
        try {
          const asset = await sdk.cma.asset.get({
            spaceId: sdk.ids.space,
            environmentId: sdk.ids.environment,
            assetId,
          });
          const sys = (asset as any)?.sys;
          next[assetId] =
            !sys?.publishedVersion
              ? 'draft'
              : typeof sys?.version === 'number' && sys.version === sys.publishedVersion + 1
                ? 'published'
                : 'changed';
        } catch {
          // ignore
        }
      });

      if (!cancelled && Object.keys(next).length) {
        setStatusAssetById((prev) => ({ ...prev, ...next }));
      }
    };

    load();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [actionableMediaRows]);

  useEffect(() => {
    let cancelled = false;

    const load = async () => {
      const ids = actionableContentRows.slice(0, 250).map((r) => r.id).filter(Boolean);
      const next: Record<string, EntityStatus> = {};
      const concurrency = 6;

      let idx = 0;
      const runners = new Array(concurrency).fill(0).map(async () => {
        while (idx < ids.length) {
          const myIdx = idx++;
          const entryId = ids[myIdx];
          if (statusEntryById[entryId] !== undefined) continue;

          try {
            const entry = await sdk.cma.entry.get({
              spaceId: sdk.ids.space,
              environmentId: sdk.ids.environment,
              entryId,
            });

            const sys = (entry as any)?.sys;
            next[entryId] =
              !sys?.publishedVersion
                ? 'draft'
                : typeof sys?.version === 'number' && sys.version === sys.publishedVersion + 1
                  ? 'published'
                  : 'changed';
          } catch {
            // ignore
          }
        }
      });

      await Promise.all(runners);

      if (!cancelled && Object.keys(next).length) {
        setStatusEntryById((prev) => ({ ...prev, ...next }));
      }
    };

    load();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [actionableContentRows]);

  const initialContentSelected = useMemo(
    () => actionableContentRows.map((r) => r.id),
    [actionableContentRows]
  );
  const initialMediaSelected = useMemo(
    () => actionableMediaRows.map((r) => r.id),
    [actionableMediaRows]
  );

  const [selectedContentIds, setSelectedContentIds] = useState<string[]>(initialContentSelected);
  const [selectedMediaIds, setSelectedMediaIds] = useState<string[]>(initialMediaSelected);

  const [isApplying, setIsApplying] = useState(false);
  const [applyStatus, setApplyStatus] = useState('');
  const [applyError, setApplyError] = useState('');
  const [applyDone, setApplyDone] = useState(false);

  type ApplySummary = {
    entries: {
      selected: number;
      updated: number;
      republished: number;
      skippedExcluded: number;
      skippedAlreadyOk: number;
      failed: number;
    };
    assets: {
      selected: number;
      updated: number;
      republished: number;
      skippedExcluded: number;
      skippedAlreadyOk: number;
      failed: number;
    };
  };

  const [applySummary, setApplySummary] = useState<ApplySummary | null>(null);

  const contentSelectableIds = useMemo(
    () => actionableContentRows.map((r) => r.id),
    [actionableContentRows]
  );
  const mediaSelectableIds = useMemo(
    () => actionableMediaRows.map((r) => r.id),
    [actionableMediaRows]
  );

  const toggleId = (list: string[], id: string) =>
    list.includes(id) ? list.filter((x) => x !== id) : [...list, id];

  const onApply = async () => {
    setApplyError('');
    setApplyDone(false);
    setApplySummary(null);

    if (!p.targetTagIds.length) {
      setApplyError('No valid Contentful Tag IDs were resolved. Nothing to apply.');
      return;
    }

    const totalSelected = selectedContentIds.length + selectedMediaIds.length;
    if (!totalSelected) {
      setApplyError('Nothing selected.');
      return;
    }

    // Lookup maps for rows (used for preflight + apply)
    const contentById = new Map(p.contentRows.map((r) => [r.id, r] as const));
    const mediaById = new Map(p.mediaRows.map((r) => [r.id, r] as const));

    const actionableContentIds = selectedContentIds.filter((id) => {
      const row = contentById.get(id);
      return row && !row.excluded && row.willAdd;
    });

    const actionableMediaIds = selectedMediaIds.filter((id) => {
      const row = mediaById.get(id);
      return row && !row.excluded && row.willAdd;
    });

    if (!actionableContentIds.length && !actionableMediaIds.length) {
      setApplyStatus('Nothing to update — all selected items already have the tag(s).');
      return;
    }

    setIsApplying(true);
    setApplyStatus('Preparing updates…');

    try {
      const concurrency = 4;

      const summary: ApplySummary = {
        entries: {
          selected: actionableContentIds.length,
          updated: 0,
          republished: 0,
          skippedExcluded: 0,
          skippedAlreadyOk: 0,
          failed: 0,
        },
        assets: {
          selected: actionableMediaIds.length,
          updated: 0,
          republished: 0,
          skippedExcluded: 0,
          skippedAlreadyOk: 0,
          failed: 0,
        },
      };

      let entryProcessed = 0;
      let assetProcessed = 0;

      if (actionableContentIds.length) {
        setApplyStatus(`Updating content (0/${actionableContentIds.length})…`);
        await promisePool(actionableContentIds, concurrency, async (entryId) => {
          // Progress and skipping
          const row = contentById.get(entryId);
          if (row?.excluded) {
            summary.entries.skippedExcluded += 1;
            entryProcessed += 1;
            setApplyStatus(`Updating content (${entryProcessed}/${actionableContentIds.length})…`);
            return;
          }
          try {
            const entry = await sdk.cma.entry.get({
              spaceId: sdk.ids.space,
              environmentId: sdk.ids.environment,
              entryId,
            });
            const existing = Array.isArray((entry as any)?.metadata?.tags) ? (entry as any).metadata.tags : [];
            const existingIds = existing.map((t: any) => t?.sys?.id).filter(Boolean);
            const mergedLinks = [...existing];
            for (const tagId of p.targetTagIds) {
              if (!existingIds.includes(tagId)) mergedLinks.push(makeTagLink(tagId));
            }
            const sys = (entry as any)?.sys;
            const isCurrentlyPublished =
              Boolean(sys?.publishedVersion) && typeof sys?.version === 'number' && sys.version === sys.publishedVersion + 1;

            const updated = await sdk.cma.entry.update(
              { spaceId: sdk.ids.space, environmentId: sdk.ids.environment, entryId },
              {
                ...(entry as any),
                metadata: { ...(entry as any).metadata, tags: mergedLinks },
              } as any
            );
            summary.entries.updated += 1;

            // Republish ONLY if the item was fully Published (not “Changed”) before we updated metadata.
            if (preservePublishState && isCurrentlyPublished) {
              await sdk.cma.entry.publish(
                { spaceId: sdk.ids.space, environmentId: sdk.ids.environment, entryId },
                updated as any
              );
              summary.entries.republished += 1;
            }
          } catch {
            // Update failed
            summary.entries.failed += 1;
          } finally {
            entryProcessed += 1;
            setApplyStatus(`Updating content (${entryProcessed}/${actionableContentIds.length})…`);
          }
        });
      }

      if (actionableMediaIds.length) {
        setApplyStatus(`Updating media (0/${actionableMediaIds.length})…`);
        await promisePool(actionableMediaIds, concurrency, async (assetId) => {
          const row = mediaById.get(assetId);
          if (row?.excluded) {
            summary.assets.skippedExcluded += 1;
            assetProcessed += 1;
            setApplyStatus(`Updating media (${assetProcessed}/${actionableMediaIds.length})…`);
            return;
          }
          try {
            const asset = await sdk.cma.asset.get({
              spaceId: sdk.ids.space,
              environmentId: sdk.ids.environment,
              assetId,
            });
            const existing = Array.isArray((asset as any)?.metadata?.tags) ? (asset as any).metadata.tags : [];
            const existingIds = existing.map((t: any) => t?.sys?.id).filter(Boolean);
            const mergedLinks = [...existing];
            for (const tagId of p.targetTagIds) {
              if (!existingIds.includes(tagId)) mergedLinks.push(makeTagLink(tagId));
            }
            const sys = (asset as any)?.sys;
            const isCurrentlyPublished =
              Boolean(sys?.publishedVersion) && typeof sys?.version === 'number' && sys.version === sys.publishedVersion + 1;

            const updated = await sdk.cma.asset.update(
              { spaceId: sdk.ids.space, environmentId: sdk.ids.environment, assetId },
              {
                ...(asset as any),
                metadata: { ...(asset as any).metadata, tags: mergedLinks },
              } as any
            );
            summary.assets.updated += 1;

            // Republish ONLY if the item was fully Published (not “Changed”) before we updated metadata.
            if (preservePublishState && isCurrentlyPublished) {
              await sdk.cma.asset.publish(
                { spaceId: sdk.ids.space, environmentId: sdk.ids.environment, assetId },
                updated as any
              );
              summary.assets.republished += 1;
            }
          } catch {
            summary.assets.failed += 1;
          } finally {
            assetProcessed += 1;
            setApplyStatus(`Updating media (${assetProcessed}/${actionableMediaIds.length})…`);
          }
        });
      }

      setApplySummary(summary);
      setApplyStatus('Done.');
      setApplyDone(true);

      const failed = summary.entries.failed + summary.assets.failed;

      if (failed > 0) {
        await Notification.error(`Tags applied with ${failed} error(s). Review the summary for details.`);
      } else {
        await Notification.success('Tags were applied successfully.');
      }
    } catch (e: any) {
      const msg = e?.message ?? String(e);
      setApplyError(msg);
      await Notification.error(msg);
    } finally {
      setIsApplying(false);
    }
  };

  return (
    <Flex
      flexDirection="column"
      gap="spacingM"
      padding="spacingM"
      style={{ width: '100%', maxWidth: '100%', height: '100%' }}
    >
      {p.missingTargetTagsByName?.length ? (
        <Note variant="warning">
          These tag names were not found as Contentful Tags (they will be ignored):{' '}
          {p.missingTargetTagsByName.join(', ')}
        </Note>
      ) : null}

      <Stack flexDirection="column" spacing="spacingXs">
        <Text fontWeight="fontWeightDemiBold">Tags to add</Text>
        {p.targetTags?.length ? (
          <Stack flexDirection="row" spacing="spacingXs" flexWrap="wrap">
            {p.targetTags.map((t) => (
              <Badge key={t.id}>{t.name}</Badge>
            ))}
          </Stack>
        ) : (
          <Paragraph>—</Paragraph>
        )}
      </Stack>


      {!hasAnythingToApply ? (
        <Note variant="primary" title="Nothing to apply">
          All linked content and media already have the required tag(s). No changes are needed.
        </Note>
      ) : null}

      {hasAnythingToApply ? (
        <Flex gap="spacingM" alignItems="flex-start" style={{ width: '100%' }}>
          {/* Content column */}
          <Stack flexDirection="column" spacing="spacingS" style={{ flex: 1, minWidth: 0 }}>
            <Stack flexDirection="row" spacing="spacingS" alignItems="center" flexWrap="wrap">
              <Text fontWeight="fontWeightDemiBold">Content</Text>
              <Button
                variant="secondary"
                size="small"
                isDisabled={isApplying || !contentSelectableIds.length}
                onClick={() => setSelectedContentIds(contentSelectableIds)}
              >
                Select all
              </Button>
              <Button
                variant="secondary"
                size="small"
                isDisabled={isApplying || !selectedContentIds.length}
                onClick={() => setSelectedContentIds([])}
              >
                Clear
              </Button>
              <Text fontColor="gray600" fontSize="fontSizeS">
                Selected: {selectedContentIds.length}
              </Text>
            </Stack>

            <Table style={{ width: '100%' }}>
              <TableHead>
                <TableRow>
                  <TableCell />
                  <TableCell>Title</TableCell>
                  <TableCell>Status</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {actionableContentRows.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={3}>
                      <Text fontColor="gray600">No content items to update.</Text>
                    </TableCell>
                  </TableRow>
                ) : (
                  actionableContentRows.map((r) => {
                    const disabled = r.excluded || isApplying;
                    const checked = selectedContentIds.includes(r.id);
                    const statusVal = statusEntryById[r.id];

                    const statusNode = r.excluded
                      ? renderStatusBadge('excluded', 'Excluded')
                      : statusVal === undefined
                        ? <StatusSkeleton />
                        : statusVal === 'published'
                          ? renderStatusBadge('published', 'Published')
                          : statusVal === 'changed'
                            ? renderStatusBadge('changed', 'Changed')
                            : renderStatusBadge('draft', 'Draft');

                    return (
                      <TableRow key={r.id}>
                        <TableCell>
                          <Checkbox
                            isChecked={checked}
                            isDisabled={disabled}
                            onChange={() => setSelectedContentIds((prev) => toggleId(prev, r.id))}
                          />
                        </TableCell>
                        <TableCell>
                          <Tooltip content={`Type: ${r.contentType}`}>
                            <TextLink
                              href={`https://app.contentful.com/spaces/${sdk.ids.space}/environments/${sdk.ids.environment}/entries/${r.id}`}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {r.title}
                            </TextLink>
                          </Tooltip>
                        </TableCell>
                        <TableCell>{statusNode}</TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
          </Stack>

          {/* Media column */}
          <Stack flexDirection="column" spacing="spacingS" style={{ flex: 1, minWidth: 0 }}>
            <Stack flexDirection="row" spacing="spacingS" alignItems="center" flexWrap="wrap">
              <Text fontWeight="fontWeightDemiBold">Media</Text>
              <Button
                variant="secondary"
                size="small"
                isDisabled={isApplying || !mediaSelectableIds.length}
                onClick={() => setSelectedMediaIds(mediaSelectableIds)}
              >
                Select all
              </Button>
              <Button
                variant="secondary"
                size="small"
                isDisabled={isApplying || !selectedMediaIds.length}
                onClick={() => setSelectedMediaIds([])}
              >
                Clear
              </Button>
              <Text fontColor="gray600" fontSize="fontSizeS">
                Selected: {selectedMediaIds.length}
              </Text>
            </Stack>

            <Table style={{ width: '100%' }}>
              <TableHead>
                <TableRow>
                  <TableCell />
                  <TableCell>Preview</TableCell>
                  <TableCell>Title</TableCell>
                  <TableCell>Status</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {actionableMediaRows.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={4}>
                      <Text fontColor="gray600">No media items to update.</Text>
                    </TableCell>
                  </TableRow>
                ) : (
                  actionableMediaRows.map((r) => {
                    const disabled = r.excluded || isApplying;
                    const checked = selectedMediaIds.includes(r.id);
                    const statusVal = statusAssetById[r.id];

                    const statusNode = r.excluded
                      ? renderStatusBadge('excluded', 'Excluded')
                      : statusVal === undefined
                        ? <StatusSkeleton />
                        : statusVal === 'published'
                          ? renderStatusBadge('published', 'Published')
                          : statusVal === 'changed'
                            ? renderStatusBadge('changed', 'Changed')
                            : renderStatusBadge('draft', 'Draft');

                    return (
                      <TableRow key={r.id}>
                        <TableCell>
                          <Checkbox
                            isChecked={checked}
                            isDisabled={disabled}
                            onChange={() => setSelectedMediaIds((prev) => toggleId(prev, r.id))}
                          />
                        </TableCell>
                        <TableCell>
                          {mediaPreviewById[r.id] ? (
                            <img
                              src={mediaPreviewById[r.id]}
                              alt=""
                              style={{ width: 40, height: 40, objectFit: 'cover', borderRadius: 4 }}
                            />
                          ) : (
                            <ImageSkeleton />
                          )}
                        </TableCell>
                        <TableCell>
                          <TextLink
                            href={`https://app.contentful.com/spaces/${sdk.ids.space}/environments/${sdk.ids.environment}/assets/${r.id}`}
                            target="_blank"
                            rel="noopener noreferrer"
                          >
                            {r.title}
                          </TextLink>
                        </TableCell>
                        <TableCell>{statusNode}</TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
          </Stack>
        </Flex>
      ) : null}

      {applyError ? <Note variant="negative">Apply failed: {applyError}</Note> : null}
      {applyDone && applySummary ? (
        <Note variant="positive" title="Apply complete">
          <Stack flexDirection="column" spacing="spacingXs">
            <Text>
              Content: updated {applySummary.entries.updated}/{applySummary.entries.selected}, republished {applySummary.entries.republished}, skipped (excluded {applySummary.entries.skippedExcluded}, already ok {applySummary.entries.skippedAlreadyOk}), update failed {applySummary.entries.failed}.
            </Text>
            <Text>
              Media: updated {applySummary.assets.updated}/{applySummary.assets.selected}, republished {applySummary.assets.republished}, skipped (excluded {applySummary.assets.skippedExcluded}, already ok {applySummary.assets.skippedAlreadyOk}), update failed {applySummary.assets.failed}.
            </Text>
          </Stack>
        </Note>
      ) : null}
      {!applyDone && applyStatus ? (
        <Note variant="primary" title="In progress">
          {applyStatus}
        </Note>
      ) : null}

      <div
        style={{
          marginTop: 'auto',
          position: 'sticky',
          bottom: 0,
          background: 'white',
          borderTop: '1px solid #e5ebed',
          paddingTop: 12,
          paddingBottom: 8,
          width: '100%',
        }}
      >
        <Flex justifyContent="space-between" alignItems="flex-start" style={{ width: '100%' }}>
          <Stack flexDirection="column" spacing="spacing2Xs" style={{ maxWidth: 520 }} alignItems="start">
            {hasAnythingToApply && !applyDone ? (
              <>
                <Switch
                  id="preserve-publish-state"
                  isChecked={preservePublishState}
                  onChange={() => setPreservePublishState((v) => !v)}
                >
                  Preserve publish state
                </Switch>
                <Text
                  fontColor="gray600"
                  fontSize="fontSizeS"
                  style={{
                    display: '-webkit-box',
                    WebkitLineClamp: 2,
                    WebkitBoxOrient: 'vertical',
                    overflow: 'hidden',
                  }}
                >
                  If enabled, items that were published before will be republished after tags are added.
                </Text>
              </>
            ) : null}
          </Stack>

          <Stack flexDirection="row" spacing="spacingS" justifyContent="flex-end" alignItems="center">
            {!hasAnythingToApply ? (
              <Button variant="primary" onClick={close}>
                Close
              </Button>
            ) : applyDone ? (
              <Button variant="primary" onClick={close}>
                Close
              </Button>
            ) : (
              <>
                <Button variant="secondary" isDisabled={isApplying} onClick={close}>
                  Cancel
                </Button>
                <Button variant="primary" isDisabled={isApplying} isLoading={isApplying} onClick={onApply}>
                  Apply tags
                </Button>
              </>
            )}
          </Stack>
        </Flex>
      </div>
    </Flex>
  );
};

export default Dialog;
