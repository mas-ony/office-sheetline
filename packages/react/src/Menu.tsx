/**
 * Pop-up menus with submenus (top menu, context menus). On phones the same items are shown in a
 * bottom sheet, where submenus open in place with a back button.
 */
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from './Icon';
import { type IconSetName, type MenuAnchor, type MenuItem, paintIconSet, placeMenu, type Preview } from '@sheetline/core/views';
import { useUI } from './context';

export interface MenuProps {
  items: MenuItem[];
  /** Position (client coordinates) of the top-left corner, or the rectangle to open next to. */
  at: MenuAnchor;
  /** `picked`: closed because an item was chosen (not dismissed). */
  onClose: (picked?: boolean) => void;
  /** Bottom sheet presentation (phones). */
  sheet?: boolean;
  title?: string;
  label?: string;
  testId?: string;
}

function MenuList({ items, onClose, at, side, depth, onBack, label, testId }: {
  items: MenuItem[];
  onClose: (picked?: boolean) => void;
  at: MenuProps['at'];
  side: 'below' | 'right';
  depth: number;
  onBack?: () => void;
  label?: string;
  testId?: string;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [active, setActive] = useState(-1);
  const [open, setOpen] = useState<{ index: number; rect: DOMRect } | null>(null);
  const hoverTimer = useRef(0);

  useLayoutEffect(() => {
    if (ref.current) placeMenu(ref.current, at, side);
  }, [at, side, items]);

  useEffect(() => {
    if (depth === 0) ref.current?.focus({ preventScroll: true });
    return () => window.clearTimeout(hoverTimer.current);
  }, [depth]);

  const enabled = items.map((it, i) => (!it.type && !it.disabled ? i : -1)).filter((i) => i >= 0);
  const activate = (i: number, el?: HTMLElement | null) => {
    const it = items[i];
    if (!it || it.type || it.disabled) return;
    if (it.submenu) {
      const btn = el ?? (ref.current?.querySelector(`[data-index="${i}"]`) as HTMLElement | null);
      if (btn) setOpen({ index: i, rect: btn.getBoundingClientRect() });
      return;
    }
    if (!it.keepOpen) onClose(true);
    it.onClick?.();
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (open) return;
    const pos = enabled.indexOf(active);
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActive(enabled[(pos + 1) % enabled.length] ?? -1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActive(enabled[(pos - 1 + enabled.length) % enabled.length] ?? -1);
    } else if (e.key === 'Enter' || e.key === ' ' || (e.key === 'ArrowRight' && items[active]?.submenu)) {
      e.preventDefault();
      if (active >= 0) activate(active);
    } else if (e.key === 'Escape' || (e.key === 'ArrowLeft' && depth > 0)) {
      e.preventDefault();
      e.stopPropagation();
      if (depth > 0 && onBack) onBack();
      else onClose();
    }
  };

  return (
    <>
      <div
        ref={ref}
        className="sh-popup"
        role="menu"
        aria-label={label}
        tabIndex={-1}
        data-testid={testId}
        onKeyDown={onKeyDown}
        onContextMenu={(e) => e.preventDefault()}
        style={{ zIndex: 100 + depth }}
      >
        {items.map((it, i) => {
          if (it.type === 'line') return <div key={i} className="sh-menu-sep" role="separator" />;
          if (it.type === 'header') return <div key={i} className="sh-menu-header">{it.title}</div>;
          if (it.type === 'gallery') return <Gallery key={i} item={it} onClose={onClose} />;
          return (
            <button
              key={i}
              type="button"
              role={it.checked !== undefined ? 'menuitemcheckbox' : 'menuitem'}
              aria-checked={it.checked}
              aria-haspopup={it.submenu ? 'menu' : undefined}
              aria-expanded={it.submenu ? open?.index === i : undefined}
              data-index={i}
              data-testid={it.testId}
              title={it.tip}
              className={`sh-menu-item${active === i || open?.index === i ? ' active' : ''}`}
              disabled={it.disabled}
              onMouseEnter={(e) => {
                setActive(i);
                window.clearTimeout(hoverTimer.current);
                const el = e.currentTarget;
                if (it.submenu && !it.disabled) hoverTimer.current = window.setTimeout(() => setOpen({ index: i, rect: el.getBoundingClientRect() }), 120);
                else if (open && open.index !== i) hoverTimer.current = window.setTimeout(() => setOpen(null), 200);
              }}
              onClick={(e) => activate(i, e.currentTarget)}
            >
              <span className="icon-slot">{it.swatch ? <PreviewView preview={it.swatch} /> : it.icon ? <Icon name={it.icon} /> : it.checked ? <Icon name="check" className="check" /> : null}</span>
              <span className="title">{it.title}</span>
              {it.shortcut && <span className="shortcut">{it.shortcut}</span>}
              {it.submenu && <Icon name="arrow_right" className="sub" />}
            </button>
          );
        })}
      </div>
      {open && items[open.index]?.submenu && (
        <MenuList
          items={items[open.index].submenu!}
          onClose={onClose}
          at={open.rect}
          side="right"
          depth={depth + 1}
          onBack={() => {
            setOpen(null);
            ref.current?.focus({ preventScroll: true });
          }}
          label={items[open.index].title}
        />
      )}
    </>
  );
}

function SheetMenu({ items, onClose, title, testId }: { items: MenuItem[]; onClose: (picked?: boolean) => void; title?: string; testId?: string }) {
  const [stack, setStack] = useState<{ title?: string; items: MenuItem[] }[]>([{ title, items }]);
  const top = stack[stack.length - 1];
  return (
    <>
      <div className="sh-sheet-backdrop" onPointerDown={(e) => { e.preventDefault(); onClose(); }} />
      <div className="sh-bottom-sheet" role="menu" aria-label={top.title} data-testid={testId}>
        <div className="sh-sheet-grip" />
        {(stack.length > 1 || top.title) && (
          <div className="sh-sheet-title">
            {stack.length > 1 ? (
              <button type="button" aria-label="Back" onClick={() => setStack(stack.slice(0, -1))}>
                <Icon name="arrow_right" className="sh-flip" />
              </button>
            ) : <span style={{ width: 8 }} />}
            <span>{top.title}</span>
          </div>
        )}
        {top.items.map((it, i) => {
          if (it.type === 'line') return <div key={i} className="sh-menu-sep" role="separator" />;
          if (it.type === 'header') return <div key={i} className="sh-menu-header">{it.title}</div>;
          if (it.type === 'gallery') return <Gallery key={i} item={it} onClose={onClose} />;
          return (
            <button
              key={i}
              type="button"
              role="menuitem"
              data-testid={it.testId}
              title={it.tip}
              className="sh-menu-item"
              disabled={it.disabled}
              onClick={() => {
                if (it.submenu) {
                  setStack([...stack, { title: it.title, items: it.submenu }]);
                  return;
                }
                if (!it.keepOpen) onClose(true);
                it.onClick?.();
              }}
            >
              <span className="icon-slot">{it.swatch ? <PreviewView preview={it.swatch} /> : it.icon ? <Icon name={it.icon} /> : it.checked ? <Icon name="check" className="check" /> : null}</span>
              <span className="title">{it.title}</span>
              {it.submenu && <Icon name="arrow_right" className="sub" />}
            </button>
          );
        })}
      </div>
    </>
  );
}

function IconSetPreview({ set }: { set: IconSetName }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    if (ref.current) paintIconSet(ref.current, set);
  }, [set]);
  return <canvas ref={ref} style={{ width: Number(set[0]) * 16, height: 16 }} aria-hidden="true" />;
}

/** A small picture of a menu item or gallery entry. */
export function PreviewView({ preview: p }: { preview: Preview }) {
  switch (p.kind) {
    case 'svg':
      return <span className="sh-svg" style={{ width: p.width, height: p.height }} dangerouslySetInnerHTML={{ __html: p.markup }} />;
    case 'color':
      return <span className="sh-menu-swatch" style={{ background: p.css }} />;
    case 'theme':
      return (
        <span className="sh-theme-swatch">
          {p.colors.map((c, i) => (
            <i key={i} style={{ background: c }} />
          ))}
        </span>
      );
    case 'font':
      return <span className="sh-font-aa" style={{ fontFamily: p.family }}>Aa</span>;
    case 'palette':
      return (
        <span className="sh-palette-row">
          {p.colors.map((c, i) => (
            <i key={i} style={{ background: c }} />
          ))}
        </span>
      );
    case 'paste':
      return (
        <span className="sh-paste-icon">
          <Icon name="content_paste" className="base" />
          {p.badge ? <Icon name={p.badge} className="badge" /> : null}
        </span>
      );
    case 'icon-set':
      return <IconSetPreview set={p.set} />;
  }
}

function Gallery({ item, onClose }: { item: MenuItem; onClose: (picked?: boolean) => void }) {
  return (
    <div className="sh-menu-gallery" role="group" aria-label={item.title}>
      {item.title && <div className="sh-menu-header">{item.title}</div>}
      <div className="grid" style={{ gridTemplateColumns: `repeat(${item.columns ?? 6}, auto)` }}>
        {(item.gallery ?? []).map((g, k) => (
          <button
            key={k}
            type="button"
            role="menuitem"
            className="sh-gallery-item"
            title={g.title}
            aria-label={g.title}
            data-testid={g.testId}
            onClick={() => {
              onClose(true);
              g.onClick();
            }}
          >
            <PreviewView preview={g.preview} />
          </button>
        ))}
      </div>
    </div>
  );
}

/** A menu shown until an item is chosen or the user clicks elsewhere. */
export function Menu({ items, at, onClose, sheet, title, label, testId }: MenuProps) {
  const { layer } = useUI();
  useEffect(() => {
    if (sheet) return;
    const down = (e: PointerEvent) => {
      const t = e.target as HTMLElement | null;
      if (t?.closest('.sh-popup')) return;
      onClose();
    };
    const blur = () => onClose();
    // Opening clicks must not close the menu immediately.
    const id = window.setTimeout(() => {
      window.addEventListener('pointerdown', down, true);
      window.addEventListener('resize', blur);
      window.addEventListener('blur', blur);
    }, 0);
    return () => {
      window.clearTimeout(id);
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('resize', blur);
      window.removeEventListener('blur', blur);
    };
  }, [onClose, sheet]);
  const content: ReactNode = sheet ? (
    <SheetMenu items={items} onClose={onClose} title={title ?? label} testId={testId} />
  ) : (
    <MenuList items={items} onClose={onClose} at={at} side="below" depth={0} label={label ?? title} testId={testId} />
  );
  return createPortal(content, layer);
}
