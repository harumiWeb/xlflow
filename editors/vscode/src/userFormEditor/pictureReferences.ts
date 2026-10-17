import { DraftState } from "./draft";

export interface PictureReferenceSources {
  /** Read live buffers again after asynchronous scans, including native history. */
  openTexts: () => Iterable<string>;
  sourceTexts: () => AsyncIterable<string>;
  /** All persisted checkpoints, including closed editors and temporary writes. */
  checkpointTexts: () => AsyncIterable<string>;
  /** A checkpoint published during a scan makes an empty result inconclusive. */
  checkpointStable?: () => boolean;
}

/** Deletion requires positive evidence that every available source is clear.
 * Filename matching deliberately over-retains assets on unrelated references.
 */
export async function pictureAssetReferenced(
  relativePath: string,
  sources: PictureReferenceSources,
): Promise<boolean> {
  const filename = relativePath.replace(/\\/g, "/").split("/").at(-1);
  if (!filename) return true;
  try {
    const liveReference = () => {
      for (const text of sources.openTexts()) if (text.includes(filename)) return true;
      return false;
    };
    if (liveReference()) return true;
    for await (const text of sources.sourceTexts()) if (text.includes(filename)) return true;
    for await (const text of sources.checkpointTexts()) {
      if (text.includes(filename)) return true;
      // Decode JSON escapes in persisted source/history strings. An incomplete
      // temporary checkpoint or unreadable recovery data cannot authorize GC.
      const checkpoint: unknown = JSON.parse(text);
      if (
        !checkpoint ||
        typeof checkpoint !== "object" ||
        !("state" in checkpoint) ||
        !("sourceUri" in checkpoint) ||
        typeof checkpoint.sourceUri !== "string" ||
        !("updatedAt" in checkpoint) ||
        !Number.isSafeInteger(checkpoint.updatedAt) ||
        (checkpoint.updatedAt as number) < 0 ||
        !("schemaVersion" in checkpoint) ||
        checkpoint.schemaVersion !== 1
      )
        return true;
      DraftState.restore(checkpoint.state);
      if (JSON.stringify(checkpoint).includes(filename)) return true;
    }
    return liveReference() || sources.checkpointStable?.() === false;
  } catch {
    return true;
  }
}
