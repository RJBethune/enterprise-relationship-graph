import { SpRest } from './SpRest';

/**
 * What the signed-in user may actually do to a SPECIFIC list.
 *
 * The web-level permission is a different question with a frequently different answer:
 * a list can carry unique permissions, so somebody who can edit the site may be
 * read-only on this one. Trusting the web-level answer means offering that person a
 * fully editable graph that fails on their first save — after they have done the work.
 */

/** SP.PermissionKind ordinals. */
const ADD_LIST_ITEMS = 2;
const EDIT_LIST_ITEMS = 3;
const MANAGE_LISTS = 12;

export interface IListRights {
  canEdit: boolean;
  canManage: boolean;
}

export const readListRights = async (
  sp: SpRest, listTitle: string
): Promise<IListRights | null> => {
  try {
    const res = await sp.get<{ High: string | number; Low: string | number }>(
      `web/lists/getbytitle('${encodeURIComponent(listTitle)}')/EffectiveBasePermissions`
    );
    // High and Low arrive as strings and can exceed 2^31. JavaScript's bitwise
    // operators coerce through ToInt32, which is precisely correct for testing one bit:
    // 4294967295 becomes -1, and -1 has every bit set.
    const high = Number(res.High) || 0;
    const low = Number(res.Low) || 0;
    const has = (kind: number): boolean => (kind < 32
      ? (low & (1 << kind)) !== 0
      : (high & (1 << (kind - 32))) !== 0);

    return {
      canEdit: has(ADD_LIST_ITEMS) && has(EDIT_LIST_ITEMS),
      canManage: has(MANAGE_LISTS)
    };
  } catch {
    // Usually "the list does not exist yet". Returning null lets the caller fall back
    // to the web-level answer rather than locking everybody out of an unprovisioned site.
    return null;
  }
};
