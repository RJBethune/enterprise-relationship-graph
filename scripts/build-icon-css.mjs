/**
 * Generate src/engine/iconAssets.ts — Font Awesome's CLASS rules as a string.
 *
 * Why this exists, since importing the stylesheet ought to be enough:
 *
 * SPFx runs every imported stylesheet through css-loader with CSS MODULES enabled,
 * which hashes class names — `.fa-solid` becomes `.fa-solid_f1fd2f8f`. That is correct
 * and desirable for a component's own styles, and fatal for an icon library, because
 * its class names ARE its API: the engine's markup says `<i class="fa-solid fa-user">`
 * and nothing will ever match it again.
 *
 * The failure is quietly asymmetric, which is what made it confusing in the field:
 * `@font-face` is not a class, so it survives mangling untouched. Icons drawn onto the
 * CANVAS kept working — the engine asks for the family by name — while every icon in
 * the DOM rendered as an empty box.
 *
 * So the stylesheet import stays (it registers the faces and makes webpack emit the
 * woff2 files with correct public paths), and the class rules are injected separately,
 * verbatim, alongside the engine's own stylesheet. `@font-face` blocks are stripped from
 * this copy precisely because the import already provides them — with resolved URLs that
 * a string constant could not produce.
 *
 *   node scripts/build-icon-css.mjs
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const source = join(root, 'node_modules/@fortawesome/fontawesome-free/css/all.min.css');

const raw = readFileSync(source, 'utf8');

// Strip @font-face — the import supplies those, with URLs webpack has resolved.
const withoutFaces = raw.replace(/@font-face\s*\{[^}]*\}/g, '');

const before = (raw.match(/@font-face/g) || []).length;
if (before === 0) { throw new Error('No @font-face blocks found; is this the right stylesheet?'); }
if (/url\(/.test(withoutFaces)) {
  throw new Error('A url() survived outside @font-face — it would not resolve from a string constant.');
}
if (!/\.fa-user:before|\.fa-user::before/.test(withoutFaces)) {
  throw new Error('Glyph mappings missing; the stylesheet layout changed.');
}

const version = JSON.parse(
  readFileSync(join(root, 'node_modules/@fortawesome/fontawesome-free/package.json'), 'utf8')
).version;

const literal = withoutFaces
  .replace(/\\/g, '\\\\')
  .replace(/`/g, '\\`')
  .replace(/\$\{/g, '\\${')
  .trim();

const out = `/* ==========================================================================
   Font Awesome ${version} — CLASS RULES ONLY (generated; do not hand-edit)
   ==========================================================================
   Produced by scripts/build-icon-css.mjs. Regenerate after changing the
   @fortawesome/fontawesome-free version.

   These are injected as plain CSS rather than imported, because SPFx's css-loader
   runs CSS Modules over imported stylesheets and hashes class names. An icon
   library's class names are its API — mangling them leaves every <i class="fa-...">
   in the DOM rendering as an empty box, while canvas-drawn glyphs carry on working
   because they reference the font family directly.

   @font-face blocks are deliberately absent: the stylesheet import in mount.ts
   provides those, with URLs webpack resolved to the emitted woff2 files.
   ========================================================================== */

export const ICON_CSS: string = \`${literal}\`;
`;

const target = join(root, 'src/engine/iconAssets.ts');
writeFileSync(target, out, 'utf8');
console.log(
  `iconAssets.ts written: Font Awesome ${version}, ${withoutFaces.length}B of class rules ` +
  `(${before} @font-face blocks stripped)`
);
