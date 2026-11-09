/** An icon of the editor's set, as an inline SVG that takes the text colour. */
import { type IconName, iconSvg } from '@sheetline/core/views';

export function Icon({ name, size, className, color }: { name: IconName; size?: number; className?: string; color?: string }) {
  return (
    <span
      className={`mi${className ? ` ${className}` : ''}`}
      aria-hidden="true"
      style={size || color ? { width: size, height: size, color } : undefined}
      // The SVG strings are bundled package assets, not user content.
      dangerouslySetInnerHTML={{ __html: iconSvg(name) }}
    />
  );
}
