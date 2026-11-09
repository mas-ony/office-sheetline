/**
 * Document themes (Page Layout > Themes, Colors, Fonts): the built-in presets and the theme part
 * written for them.
 */
import type { Theme, ThemeSlot } from './model/color';
import { DEFAULT_THEME_XML } from './xlsx/theme';

export interface ColorScheme {
  name: string;
  colors: Record<ThemeSlot, string>;
}

export interface FontScheme {
  name: string;
  major: string;
  minor: string;
}

export interface ThemePreset {
  name: string;
  colors: string;
  fonts: string;
}

const scheme = (name: string, dk2: string, lt2: string, accents: string[], hlink: string, folHlink: string): ColorScheme => ({
  name,
  colors: {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2,
    lt2,
    accent1: accents[0],
    accent2: accents[1],
    accent3: accents[2],
    accent4: accents[3],
    accent5: accents[4],
    accent6: accents[5],
    hlink,
    folHlink,
  },
});

/** Colour schemes in the order of the Colors gallery. */
export const COLOR_SCHEMES: ColorScheme[] = [
  scheme('Office', '44546A', 'E7E6E6', ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'], '0563C1', '954F72'),
  scheme('Office 2007', '1F497D', 'EEECE1', ['4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'], '0000FF', '800080'),
  scheme('Grayscale', '000000', 'F8F8F8', ['DDDDDD', 'B2B2B2', '969696', '808080', '5F5F5F', '4D4D4D'], '5F5F5F', '919191'),
  scheme('Apex', '69676D', 'C9C2D1', ['CEB966', '9CB084', '6BB1C9', '6585CF', '7E6BC9', 'A379BB'], '410082', '932968'),
  scheme('Aspect', '323232', 'E3DED1', ['F07F09', '9F2936', '1B587C', '4E8542', '604878', 'C19859'], '6B9F25', 'B26B02'),
  scheme('Civic', '646B86', 'C5D1D7', ['D16349', 'CCB400', '8CADAE', '8C7B70', '8FB08C', 'D19049'], '00A3D6', '694F07'),
  scheme('Concourse', '464646', 'DEF5FA', ['2DA2BF', 'DA1F28', 'EB641B', '39639D', '474B78', '7D3C4A'], 'FF8119', '44B9E8'),
  scheme('Equity', '696464', 'E9E5DC', ['D34817', '9B2D1F', 'A28E6A', '956251', '918485', '855D5D'], 'CC9900', '96A9A9'),
  scheme('Flow', '04617B', 'DBF5F9', ['0F6FC6', '009DD9', '0BD0D9', '10CF9B', '7CCA62', 'A5C249'], 'E2D700', '85DFD0'),
  scheme('Foundry', '676A55', 'EAEBDE', ['72A376', 'B0CCB0', 'A8CDD7', 'C0BEAF', 'CEC597', 'E8B7B7'], 'DB5353', '903638'),
  scheme('Median', '775F55', 'EBDDC3', ['94B6D2', 'DD8047', 'A5AB81', 'D8B25C', '7BA79D', '968C8C'], 'F7B615', '704404'),
  scheme('Metro', '4E5B6F', 'D6ECFF', ['7FD13B', 'EA157A', 'FEB80A', '00ADDC', '738AC8', '1AB39F'], 'EB8803', '5F7791'),
  scheme('Module', '5A6378', 'D4D4D6', ['F0AD00', '60B5CC', 'E66C7D', '6BB76D', 'E88651', 'C64847'], '168BBA', '680000'),
  scheme('Opulent', 'B13F9A', 'F4E7ED', ['B83D68', 'AC66BB', 'DE6C36', 'F9B639', 'CF6DA4', 'FA8D3D'], 'FFDE66', 'D490C5'),
  scheme('Oriel', '575F6D', 'FFF39D', ['FE8637', '7598D9', 'B32C16', 'F5CD2D', 'AEBAD5', '777C84'], 'D2611C', '3B435B'),
  scheme('Origin', '464653', 'DDE9EC', ['727CA3', '9FB8CD', 'D2DA7A', 'FADA7A', 'B88472', '8E736A'], 'B292CA', '6B5680'),
  scheme('Paper', '444D26', 'FEFAC9', ['A5B592', 'F3A447', 'E7BC29', 'D092A7', '9C85C0', '809EC2'], '8E58B6', '7F6F6F'),
  scheme('Solstice', '4F271C', 'E7DEC9', ['3891A7', 'FEB80A', 'C32D2E', '84AA33', '964305', '475A8D'], '8DC765', 'AA8A14'),
  scheme('Technic', '3B3B3B', 'D4D2D0', ['6EA0B0', 'CCAF0A', '8D89A4', '748560', '9E9273', '7E848D'], '00C8C3', 'A116E0'),
  scheme('Trek', '4E3B30', 'FBEEC9', ['F0A22E', 'A5644E', 'B58B80', 'C3986D', 'A19574', 'C17529'], 'AD1F1F', 'FFC42F'),
  scheme('Urban', '424456', 'DEDEDE', ['53548A', '438086', 'A04DA3', 'C4652D', '8B5D3D', '5C92B5'], '67AFBD', 'C2A874'),
  scheme('Verve', '666666', 'D2D2D2', ['FF388C', 'E40059', '9C007F', '68007F', '005BD3', '00349E'], '17BBFD', 'FF79C2'),
];

export const FONT_SCHEMES: FontScheme[] = [
  { name: 'Office', major: 'Calibri Light', minor: 'Calibri' },
  { name: 'Office 2007', major: 'Cambria', minor: 'Calibri' },
  { name: 'Apex', major: 'Lucida Sans', minor: 'Book Antiqua' },
  { name: 'Aspect', major: 'Verdana', minor: 'Verdana' },
  { name: 'Civic', major: 'Georgia', minor: 'Georgia' },
  { name: 'Concourse', major: 'Lucida Sans Unicode', minor: 'Lucida Sans Unicode' },
  { name: 'Equity', major: 'Franklin Gothic Book', minor: 'Perpetua' },
  { name: 'Flow', major: 'Calibri', minor: 'Constantia' },
  { name: 'Foundry', major: 'Rockwell', minor: 'Rockwell' },
  { name: 'Median', major: 'Tw Cen MT', minor: 'Tw Cen MT' },
  { name: 'Metro', major: 'Consolas', minor: 'Corbel' },
  { name: 'Module', major: 'Corbel', minor: 'Corbel' },
  { name: 'Opulent', major: 'Trebuchet MS', minor: 'Trebuchet MS' },
  { name: 'Oriel', major: 'Century Schoolbook', minor: 'Century Schoolbook' },
  { name: 'Origin', major: 'Bookman Old Style', minor: 'Gill Sans MT' },
  { name: 'Paper', major: 'Constantia', minor: 'Constantia' },
  { name: 'Solstice', major: 'Gill Sans MT', minor: 'Gill Sans MT' },
  { name: 'Technic', major: 'Franklin Gothic Book', minor: 'Arial' },
  { name: 'Trek', major: 'Franklin Gothic Medium', minor: 'Franklin Gothic Book' },
  { name: 'Urban', major: 'Trebuchet MS', minor: 'Georgia' },
  { name: 'Verve', major: 'Century Gothic', minor: 'Century Gothic' },
];

/** Themes of the Themes gallery: a colour scheme with a font scheme. */
export const THEME_PRESETS: ThemePreset[] = [
  { name: 'Office', colors: 'Office', fonts: 'Office' },
  { name: 'Office 2007', colors: 'Office 2007', fonts: 'Office 2007' },
  ...COLOR_SCHEMES.slice(3).map((c) => ({ name: c.name, colors: c.name, fonts: c.name })),
];

export function colorScheme(name: string): ColorScheme | undefined {
  return COLOR_SCHEMES.find((s) => s.name === name);
}

export function fontScheme(name: string): FontScheme | undefined {
  return FONT_SCHEMES.find((s) => s.name === name);
}

export const THEME_SLOT_LABELS: { slot: ThemeSlot; label: string }[] = [
  { slot: 'dk1', label: 'Text/Background - Dark 1' },
  { slot: 'lt1', label: 'Text/Background - Light 1' },
  { slot: 'dk2', label: 'Text/Background - Dark 2' },
  { slot: 'lt2', label: 'Text/Background - Light 2' },
  { slot: 'accent1', label: 'Accent 1' },
  { slot: 'accent2', label: 'Accent 2' },
  { slot: 'accent3', label: 'Accent 3' },
  { slot: 'accent4', label: 'Accent 4' },
  { slot: 'accent5', label: 'Accent 5' },
  { slot: 'accent6', label: 'Accent 6' },
  { slot: 'hlink', label: 'Hyperlink' },
  { slot: 'folHlink', label: 'Followed Hyperlink' },
];

const escapeAttr = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/**
 * The theme part for a theme: the colour and font schemes of `base` (the file's own theme part,
 * or the default one) replaced, its formatting scheme kept.
 */
export function themeXml(theme: Theme, name: string, base: string = DEFAULT_THEME_XML): string {
  const slots: ThemeSlot[] = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const clr = `<a:clrScheme name="${escapeAttr(name)}">${slots
    .map((s) => {
      const v = theme.colors[s];
      if (s === 'dk1' && v === '000000') return '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
      if (s === 'lt1' && v === 'FFFFFF') return '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
      return `<a:${s}><a:srgbClr val="${v}"/></a:${s}>`;
    })
    .join('')}</a:clrScheme>`;
  let out = base.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, clr);
  out = out.replace(/(<a:majorFont>\s*)<a:latin\b[^>]*\/>/, `$1<a:latin typeface="${escapeAttr(theme.majorFont)}"/>`);
  out = out.replace(/(<a:minorFont>\s*)<a:latin\b[^>]*\/>/, `$1<a:latin typeface="${escapeAttr(theme.minorFont)}"/>`);
  out = out.replace(/(<a:theme\b[^>]*\bname=")[^"]*(")/, `$1${escapeAttr(name === 'Office' ? 'Office Theme' : name)}$2`);
  return out;
}
