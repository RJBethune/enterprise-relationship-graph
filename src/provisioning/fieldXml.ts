import { IExpectedField } from './schema';

/**
 * CAML SchemaXml for one column.
 *
 * Dialect carried over from the Task Tracker provisioner, which is tenant-proven:
 *  - Name + StaticName pin the INTERNAL name. Created through the API with the
 *    AddFieldInternalNameHint option, there is no rename trap — the internal name is
 *    exactly what is asked for, not a mangled version of the display name.
 *  - Creation carries EVERYTHING (index, uniqueness, default, choices), so a new
 *    column costs one round trip and needs no fix-up pass.
 *  - Note is ALWAYS plain text. Enhanced rich text HTML-encodes content, which would
 *    corrupt a graph payload on the first round trip.
 */

const escapeXml = (v: string): string => v
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

/** Replaced by the executor with the target list's GUID once that list exists. */
export const lookupToken = (list: string): string => `{{LOOKUP:${list}}}`;

export const fieldSchemaXml = (f: IExpectedField): string => {
  const type = f.types[0];
  const common =
    `DisplayName="${escapeXml(f.display || f.internal)}" Name="${escapeXml(f.internal)}" StaticName="${escapeXml(f.internal)}"` +
    (f.required ? ' Required="TRUE"' : '') +
    (f.indexed || f.unique ? ' Indexed="TRUE"' : '') +
    (f.unique ? ' EnforceUniqueValues="TRUE"' : '') +
    (f.description ? ` Description="${escapeXml(f.description)}"` : '');
  const defaultXml = f.defaultValue !== undefined ? `<Default>${escapeXml(f.defaultValue)}</Default>` : '';

  switch (type) {
    case 'Text':
      return `<Field Type="Text" ${common} MaxLength="255">${defaultXml}</Field>`;
    case 'Note':
      // NumLines is cosmetic; RichText FALSE is load-bearing.
      return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" />`;
    case 'Number':
      return `<Field Type="Number" ${common}>${defaultXml}</Field>`;
    case 'Boolean':
      return `<Field Type="Boolean" ${common}>${defaultXml || '<Default>0</Default>'}</Field>`;
    case 'DateTime':
      return `<Field Type="DateTime" ${common} Format="DateTime" />`;
    case 'Choice': {
      const choices = (f.choices || []).map((c) => `<CHOICE>${escapeXml(c)}</CHOICE>`).join('');
      return `<Field Type="Choice" ${common} Format="Dropdown">${defaultXml}<CHOICES>${choices}</CHOICES></Field>`;
    }
    case 'Lookup': {
      const target = f.lookup ? f.lookup.list : '';
      const behavior = f.lookup && f.lookup.behavior === 'Cascade'
        ? ' RelationshipDeleteBehavior="Cascade"'
        : '';
      return `<Field Type="Lookup" ${common} List="${lookupToken(target)}" ShowField="Title"${behavior} />`;
    }
    default:
      // A declared type this builder cannot express is a programming error, not a
      // site condition — fail loudly at the seam rather than writing bad CAML.
      throw new Error(`fieldSchemaXml: unsupported field type '${String(type)}' for ${f.internal}`);
  }
};
