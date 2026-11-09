/**
 * The ribbon: File button, tabs (with contextual tabs), and the active tab's groups. Groups that do
 * not fit collapse into group buttons (from the right), as on phones where every group is a
 * button opening a bottom sheet. The toolbar layout shows one row of small buttons instead.
 */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { Icon } from '../Icon';
import { Menu } from '../Menu';
import { type ColorSpec, type ComboOption, type Command, commandChecked, commandDisabled, commandTip, type Control, type Ctx, type Group, type MenuItem, resolveColor, runRibbonCommand, type Tab, toolbarControls } from '@sheetline/core/views';
import { RibbonPanel, RibbonView } from './Panels';
import { ThemePalette } from '../ThemePalette';
import { useUI } from '../context';

// -----------------------------------------------------------------------------------------------
// Drop-down surface

export function Dropdown({ anchor, onClose, children, className, testId, sheet, title }: {
  anchor: DOMRect;
  onClose: () => void;
  children: ReactNode;
  className?: string;
  testId?: string;
  /** Bottom sheet presentation (phones). */
  sheet?: boolean;
  title?: string;
}) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el || sheet) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    let x = anchor.left;
    let y = anchor.bottom + 1;
    if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
    if (y + h > window.innerHeight - 4) y = Math.max(4, Math.min(anchor.top - h - 1, window.innerHeight - h - 4));
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
  }, [anchor, sheet]);
  useEffect(() => {
    const down = (e: PointerEvent) => {
      const t = e.target as HTMLElement | null;
      if (!t || ref.current?.contains(t) || t.closest('.sh-color, .sh-popup, .sh-bottom-sheet, .sh-modal-layer, .sh-dropdown.nested')) return;
      onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      }
    };
    const id = window.setTimeout(() => {
      window.addEventListener('pointerdown', down, true);
      window.addEventListener('keydown', key);
    }, 0);
    return () => {
      window.clearTimeout(id);
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key);
    };
  }, [onClose]);
  if (sheet) {
    return createPortal(
      <>
        <div className="sh-sheet-backdrop" onPointerDown={(e) => { e.preventDefault(); onClose(); }} />
        <div ref={ref} className={`sh-bottom-sheet sh-sheet-panel${className ? ` ${className}` : ''}`} data-testid={testId} role="dialog" aria-label={title}>
          <div className="sh-sheet-grip" />
          {title && <div className="sh-sheet-title"><span style={{ width: 8 }} /><span>{title}</span></div>}
          {children}
        </div>
      </>,
      layer,
    );
  }
  return createPortal(
    <div ref={ref} className={`sh-dropdown${className ? ` ${className}` : ''}`} data-testid={testId} onContextMenu={(e) => e.preventDefault()}>
      {children}
    </div>,
    layer,
  );
}

// -----------------------------------------------------------------------------------------------
// Controls

type Open = { id: string; rect: DOMRect; kind: 'menu' | 'panel' | 'group' | 'color' | 'more' } | null;

interface ControlEnv {
  ctx: Ctx;
  open: Open;
  setOpen: (o: Open) => void;
  /** Toolbar layout: large buttons render as small ones. */
  compact: boolean;
  phone: boolean;
  /** Closes the popup holding the control once a command ran (toggles keep it open). */
  done?: () => void;
}

const keepFocus = (e: React.MouseEvent) => e.preventDefault();

function CommandButton({ cmd, env, size, split, iconOnly }: { cmd: Command; env: ControlEnv; size: 'large' | 'small'; split?: boolean; iconOnly?: boolean }) {
  const { ctx, open, setOpen } = env;
  const ref = useRef<HTMLDivElement>(null);
  const hasDrop = !!(cmd.menu || cmd.panel);
  const disabled = commandDisabled(cmd, ctx) || (ctx.state.readOnly && !cmd.id.startsWith('view') && !cmd.id.startsWith('file'));
  const checked = commandChecked(cmd, ctx);
  const opened = open?.id === cmd.id;
  const openDrop = () => {
    if (opened) {
      setOpen(null);
      return;
    }
    setOpen({ id: cmd.id, rect: ref.current!.getBoundingClientRect(), kind: cmd.panel ? 'panel' : 'menu' });
  };
  const onMain = () => {
    if (disabled) return;
    if (hasDrop && !(split && cmd.run)) openDrop();
    else {
      runRibbonCommand(cmd, ctx);
      if (!cmd.checked) env.done?.();
    }
  };
  const icon = cmd.icon ? <Icon name={cmd.icon} color={disabled ? undefined : cmd.tint} className={size === 'large' && !env.compact ? 'large' : undefined} /> : null;
  const showLabel = !iconOnly && !(env.compact && size === 'small' && !!cmd.icon && iconOnly !== false);
  const cls = `sh-btn ${env.compact ? 'small' : size}${checked ? ' checked' : ''}${opened ? ' open' : ''}${split && hasDrop && cmd.run ? ' split' : ''}`;
  return (
    <div ref={ref} className={cls} title={commandTip(cmd)}>
      <button
        type="button"
        className="sh-btn-main"
        aria-label={cmd.label}
        aria-pressed={cmd.checked ? checked : undefined}
        aria-haspopup={hasDrop && !(split && cmd.run) ? 'menu' : undefined}
        aria-expanded={hasDrop && !(split && cmd.run) ? opened : undefined}
        disabled={disabled}
        data-testid={`rb-${cmd.id}`}
        onMouseDown={keepFocus}
        onClick={onMain}
      >
        {icon}
        {(size === 'large' && !env.compact) ? (
          // A large split button shows its label on the arrow part.
          split && hasDrop && cmd.run ? null : (
            <span className="sh-btn-label">
              {cmd.label}
              {hasDrop && <Icon name="arrow_drop_down" className="caret" />}
            </span>
          )
        ) : showLabel || !icon ? <span className="sh-btn-label">{cmd.label}</span> : null}
        {(size === 'small' || env.compact) && hasDrop && !(split && cmd.run) && <Icon name="arrow_drop_down" className="caret" />}
      </button>
      {split && hasDrop && cmd.run && (
        <button
          type="button"
          className="sh-btn-arrow"
          aria-label={`${cmd.label} options`}
          aria-haspopup="menu"
          aria-expanded={opened}
          disabled={disabled}
          data-testid={`rb-${cmd.id}-more`}
          onMouseDown={keepFocus}
          onClick={openDrop}
        >
          {size === 'large' && !env.compact ? <span className="sh-btn-label">{cmd.label}</span> : null}
          <Icon name="arrow_drop_down" className="caret" />
        </button>
      )}
      {opened && open.kind === 'menu' && cmd.menu && (
        <Menu
          items={cmd.menu(ctx)}
          at={open.rect}
          sheet={env.phone}
          title={cmd.label}
          testId={`rb-${cmd.id}-menu`}
          onClose={(picked) => {
            setOpen(null);
            if (picked) env.done?.();
            ctx.grid?.focusIdle();
          }}
        />
      )}
      {opened && open.kind === 'panel' && cmd.panel && (
        <Dropdown anchor={open.rect} sheet={env.phone} title={cmd.label} testId={`rb-${cmd.id}-panel`} onClose={() => setOpen(null)}>
          <RibbonPanel spec={cmd.panel} ctx={ctx} close={() => { setOpen(null); env.done?.(); ctx.grid?.focusIdle(); }} />
        </Dropdown>
      )}
    </div>
  );
}

function ComboControl({ c, env }: { c: Extract<Control, { type: 'combo' }>; env: ControlEnv }) {
  const { ctx } = env;
  const value = c.value(ctx);
  const [text, setText] = useState(value);
  const [focused, setFocused] = useState(false);
  const wrap = useRef<HTMLDivElement>(null);
  const opened = env.open?.id === c.id;
  useEffect(() => {
    if (!focused) setText(value);
  }, [value, focused]);
  const disabled = c.disabled?.(ctx) || ctx.state.readOnly;
  const commit = (v: string) => {
    setFocused(false);
    if (v && v !== value) c.onChange(ctx, v);
    ctx.grid?.focusIdle();
  };
  const options: ComboOption[] = opened ? c.options(ctx) : [];
  return (
    <div ref={wrap} className={`sh-combo${opened ? ' open' : ''}`} style={{ width: c.width }} title={c.label}>
      <input
        aria-label={c.label}
        data-testid={`rb-${c.id}`}
        value={text}
        disabled={disabled}
        readOnly={!c.editable}
        spellCheck={false}
        onFocus={(e) => {
          setFocused(true);
          e.target.select();
        }}
        onBlur={() => setFocused(false)}
        onChange={(e) => setText(e.target.value)}
        onMouseDown={(e) => {
          if (!c.editable) {
            e.preventDefault();
            env.setOpen(opened ? null : { id: c.id, rect: wrap.current!.getBoundingClientRect(), kind: 'menu' });
          }
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            commit(text.trim());
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setText(value);
            ctx.grid?.focusIdle();
          }
        }}
      />
      <button
        type="button"
        className="sh-combo-arrow"
        aria-label={`${c.label} list`}
        disabled={disabled}
        data-testid={`rb-${c.id}-list`}
        onMouseDown={keepFocus}
        onClick={() => env.setOpen(opened ? null : { id: c.id, rect: wrap.current!.getBoundingClientRect(), kind: 'menu' })}
      >
        <Icon name="arrow_drop_down" />
      </button>
      {opened && (
        <Dropdown anchor={env.open!.rect} sheet={env.phone} title={c.label} className="sh-combo-list" testId={`rb-${c.id}-options`} onClose={() => env.setOpen(null)}>
          <div role="listbox" aria-label={c.label}>
            {options.map((o) => (
              <button
                key={o.value}
                type="button"
                role="option"
                aria-selected={o.value === value}
                className={`sh-combo-item${o.value === value ? ' selected' : ''}`}
                style={o.font ? { fontFamily: o.font } : undefined}
                onMouseDown={keepFocus}
                onClick={() => {
                  env.setOpen(null);
                  commit(o.value);
                }}
              >
                {o.label ?? o.value}
              </button>
            ))}
          </div>
        </Dropdown>
      )}
    </div>
  );
}

function ColorControl({ c, env }: { c: Extract<Control, { type: 'color' }>; env: ControlEnv }) {
  const { ctx } = env;
  const [last, setLast] = useState<{ spec: ColorSpec | null; css: string }>({ spec: null, css: c.initial });
  const ref = useRef<HTMLDivElement>(null);
  const opened = env.open?.id === c.cmd.id;
  const disabled = commandDisabled(c.cmd, ctx) || ctx.state.readOnly;
  const apply = (spec: ColorSpec | null) => {
    env.setOpen(null);
    if (ctx.grid?.isEditing()) ctx.grid.commitEdit();
    c.apply(ctx, spec);
    const css = spec ? `#${resolveColor(spec, ctx.doc.wb.theme, '000000')}` : c.noneLabel === 'No Fill' ? 'transparent' : c.initial;
    setLast({ spec, css });
    ctx.grid?.focusIdle();
  };
  return (
    <div ref={ref} className={`sh-btn small split color${opened ? ' open' : ''}`} title={c.cmd.tip ?? c.cmd.label}>
      <button
        type="button"
        className="sh-btn-main"
        aria-label={c.cmd.label}
        disabled={disabled}
        data-testid={`rb-${c.cmd.id}`}
        onMouseDown={keepFocus}
        onClick={() => {
          if (ctx.grid?.isEditing()) ctx.grid.commitEdit();
          c.apply(ctx, last.spec ?? (c.noneLabel === 'No Fill' ? { rgb: 'FFFFFF00' } : null));
        }}
      >
        {c.cmd.icon && <Icon name={c.cmd.icon} />}
        <span className="sh-color-bar" style={{ background: last.css }} />
      </button>
      <button
        type="button"
        className="sh-btn-arrow"
        aria-label={`${c.cmd.label} palette`}
        disabled={disabled}
        data-testid={`rb-${c.cmd.id}-more`}
        onMouseDown={keepFocus}
        onClick={() => env.setOpen(opened ? null : { id: c.cmd.id, rect: ref.current!.getBoundingClientRect(), kind: 'color' })}
      >
        <Icon name="arrow_drop_down" className="caret" />
      </button>
      {opened && (
        <Dropdown anchor={env.open!.rect} sheet={env.phone} title={c.cmd.label} testId={`rb-${c.cmd.id}-palette`} onClose={() => env.setOpen(null)}>
          <ThemePalette theme={ctx.doc.wb.theme} noneLabel={c.noneLabel} noneSwatch={c.noneLabel === 'Automatic' ? '#000000' : undefined} onPick={apply} testId={`palette-${c.cmd.id}`} />
        </Dropdown>
      )}
    </div>
  );
}

function CheckControl({ cmd, env }: { cmd: Command; env: ControlEnv }) {
  const { ctx } = env;
  const checked = commandChecked(cmd, ctx);
  const disabled = commandDisabled(cmd, ctx);
  return (
    <label className={`sh-check${disabled ? ' disabled' : ''}`} title={commandTip(cmd)}>
      <input
        type="checkbox"
        checked={checked}
        disabled={disabled}
        data-testid={`rb-${cmd.id}`}
        aria-label={cmd.tip ?? cmd.label}
        onChange={() => runRibbonCommand(cmd, ctx)}
      />
      <span>{cmd.label}</span>
    </label>
  );
}

function ControlView({ c, env }: { c: Control; env: ControlEnv }) {
  switch (c.type) {
    case 'large':
      return <CommandButton cmd={c.cmd} env={env} size="large" split={c.split} />;
    case 'small':
      return <CommandButton cmd={c.cmd} env={env} size="small" split={c.split} iconOnly={c.iconOnly} />;
    case 'stack':
      return (
        <div className="sh-stack">
          {c.items.map((x, i) => <ControlView key={i} c={x} env={env} />)}
        </div>
      );
    case 'row':
      return (
        <div className="sh-row">
          {c.items.map((x, i) => <ControlView key={i} c={x} env={env} />)}
        </div>
      );
    case 'combo':
      return <ComboControl c={c} env={env} />;
    case 'color':
      return <ColorControl c={c} env={env} />;
    case 'check':
      return <CheckControl cmd={c.cmd} env={env} />;
    case 'view':
      return <RibbonView spec={c.view} ctx={env.ctx} />;
    default:
      return null;
  }
}

// -----------------------------------------------------------------------------------------------
// Groups

function GroupView({ g, env, collapsed }: { g: Group; env: ControlEnv; collapsed: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  const opened = env.open?.id === `group:${g.id}`;
  if (collapsed) {
    return (
      <div className="sh-group collapsed" data-group={g.id} ref={ref}>
        <div className={`sh-btn large${opened ? ' open' : ''}`} title={g.label}>
          <button
            type="button"
            className="sh-btn-main"
            aria-label={g.label}
            aria-haspopup="dialog"
            aria-expanded={opened}
            data-testid={`rb-group-${g.id}`}
            onMouseDown={keepFocus}
            onClick={() => env.setOpen(opened ? null : { id: `group:${g.id}`, rect: ref.current!.getBoundingClientRect(), kind: 'group' })}
          >
            <Icon name={g.icon} className="large" />
            <span className="sh-btn-label">{g.label}<Icon name="arrow_drop_down" className="caret" /></span>
          </button>
        </div>
        {opened && (
          <GroupPopup g={g} env={env} />
        )}
      </div>
    );
  }
  return (
    <div className="sh-group" data-group={g.id} ref={ref}>
      <div className="sh-group-body">
        {g.controls.map((c, i) => <ControlView key={i} c={c} env={env} />)}
      </div>
      <div className="sh-group-label">
        <span>{g.label}</span>
        {g.launcher && (
          <button
            type="button"
            className="sh-launcher"
            title={g.launcher.tip ?? g.launcher.label}
            aria-label={g.launcher.tip ?? g.launcher.label}
            data-testid={`rb-${g.launcher.id}`}
            onMouseDown={keepFocus}
            onClick={() => runRibbonCommand(g.launcher!, env.ctx)}
          >
            <svg viewBox="0 0 8 8" width="8" height="8" aria-hidden="true"><path d="M0 0h1v7h7v1H0zm3 2 1-1 3 3V2h1v4H4V5h2z" fill="currentColor" /></svg>
          </button>
        )}
      </div>
    </div>
  );
}

/** The controls of a collapsed group, in a drop-down (desktop) or a bottom sheet (phone). */
function GroupPopup({ g, env }: { g: Group; env: ControlEnv }) {
  // Nested drop-downs open on top of this popup; keep their state separate.
  const [inner, setInner] = useState<Open>(null);
  const innerEnv: ControlEnv = { ...env, open: inner, setOpen: setInner, compact: false, done: () => env.setOpen(null) };
  return (
    <Dropdown anchor={env.open!.rect} sheet={env.phone} title={g.label} className="sh-group-popup" testId={`rb-group-${g.id}-panel`} onClose={() => { if (!inner) env.setOpen(null); }}>
      <div className={`sh-group-popup-body${env.phone ? ' touch' : ''}`}>
        {g.controls.map((c, i) => <ControlView key={i} c={c} env={innerEnv} />)}
        {g.launcher && (
          <button type="button" className="sh-group-more" onClick={() => { env.setOpen(null); runRibbonCommand(g.launcher!, env.ctx); }}>
            {g.launcher.label}
          </button>
        )}
      </div>
    </Dropdown>
  );
}

// -----------------------------------------------------------------------------------------------
// Toolbar layout (one row)

function ToolbarRow({ tab, env }: { tab: Tab; env: ControlEnv }) {
  const { ctx } = env;
  const controls = useMemo(() => toolbarControls(tab, ctx), [tab, ctx]);
  const row = useRef<HTMLDivElement>(null);
  const [visible, setVisible] = useState(controls.length);
  useLayoutEffect(() => {
    setVisible(controls.length);
  }, [controls]);
  useLayoutEffect(() => {
    const el = row.current;
    if (!el || env.phone) return;
    if (el.scrollWidth > el.clientWidth + 1 && visible > 3) setVisible((v) => v - 1);
  });
  useEffect(() => {
    const fn = () => setVisible(controls.length);
    window.addEventListener('resize', fn);
    return () => window.removeEventListener('resize', fn);
  }, [controls]);
  const hidden = controls.slice(visible);
  const moreRef = useRef<HTMLButtonElement>(null);
  const opened = env.open?.id === 'toolbar-more';
  return (
    <div className="sh-toolbar-row" ref={row}>
      {controls.slice(0, visible).map((c, i) => <ControlView key={i} c={c} env={{ ...env, compact: true }} />)}
      {hidden.length > 0 && (
        <div className="sh-btn small">
          <button ref={moreRef} type="button" className="sh-btn-main" aria-label="More commands" data-testid="rb-toolbar-more" onMouseDown={keepFocus} onClick={() => env.setOpen(opened ? null : { id: 'toolbar-more', rect: moreRef.current!.getBoundingClientRect(), kind: 'more' })}>
            <Icon name="more_horiz" />
          </button>
          {opened && (
            <OverflowPopup controls={hidden} env={env} />
          )}
        </div>
      )}
    </div>
  );
}

function OverflowPopup({ controls, env }: { controls: Control[]; env: ControlEnv }) {
  const [inner, setInner] = useState<Open>(null);
  return (
    <Dropdown anchor={env.open!.rect} sheet={env.phone} className="sh-group-popup" testId="rb-toolbar-more-panel" onClose={() => { if (!inner) env.setOpen(null); }}>
      <div className="sh-group-popup-body column">
        {controls.map((c, i) => <ControlView key={i} c={c} env={{ ...env, open: inner, setOpen: setInner, compact: false, done: () => env.setOpen(null) }} />)}
      </div>
    </Dropdown>
  );
}

// -----------------------------------------------------------------------------------------------
// The ribbon

export interface RibbonProps {
  tabs: Tab[];
  active: string;
  onTab: (id: string) => void;
  onFile: () => void;
  layout: 'ribbon' | 'toolbar';
  collapsed: boolean;
  onCollapse: (v: boolean) => void;
  /** Right side of the tab row (theme and layout switches). */
  extra?: ReactNode;
}

export function Ribbon({ tabs, active, onTab, onFile, layout, collapsed, onCollapse, extra }: RibbonProps) {
  const ctx = useUI();
  const phone = ctx.phone;
  const [open, setOpen] = useState<Open>(null);
  const visibleTabs = tabs.filter((t) => !t.visible || t.visible(ctx));
  const tab = visibleTabs.find((t) => t.id === active) ?? visibleTabs.find((t) => t.id === 'home') ?? visibleTabs[0];
  const groups = useMemo(() => tab.groups(ctx), [tab, ctx]);
  const env: ControlEnv = { ctx, open, setOpen, compact: false, phone };
  const content = useRef<HTMLDivElement>(null);
  const [collapsedCount, setCollapsedCount] = useState(phone ? groups.length : 0);
  const [peek, setPeek] = useState(false);

  // Adaptive layout: collapse groups from the right until the tab fits.
  useLayoutEffect(() => {
    setCollapsedCount(phone ? groups.length : 0);
  }, [tab.id, phone, groups.length]);
  useLayoutEffect(() => {
    const el = content.current;
    if (!el || phone || layout !== 'ribbon') return;
    if (el.scrollWidth > el.clientWidth + 1 && collapsedCount < groups.length) setCollapsedCount((n) => n + 1);
  });
  useEffect(() => {
    const fn = () => setCollapsedCount(phone ? groups.length : 0);
    window.addEventListener('resize', fn);
    return () => window.removeEventListener('resize', fn);
  }, [phone, groups.length]);
  useEffect(() => setOpen(null), [tab.id, layout]);

  const showContent = !collapsed || peek;
  return (
    <div className={`sh-ribbon ${layout}${collapsed ? ' collapsed' : ''}${phone ? ' phone' : ''}`} data-testid="ribbon" role="toolbar" aria-label="Ribbon">
      <div className="sh-tabrow">
        <button type="button" className="sh-file-btn" data-testid="file-button" onMouseDown={keepFocus} onClick={onFile}>File</button>
        <div className="sh-tabs" role="tablist">
          {visibleTabs.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={t.id === tab.id}
              className={`sh-tab${t.id === tab.id ? ' active' : ''}${t.contextual ? ' contextual' : ''}`}
              data-testid={`tab-${t.id}`}
              onMouseDown={keepFocus}
              onClick={() => {
                if (collapsed) setPeek(t.id !== tab.id || !peek);
                onTab(t.id);
              }}
              onDoubleClick={() => onCollapse(!collapsed)}
            >
              {t.label}
            </button>
          ))}
        </div>
        <div className="sh-tabrow-extra">{extra}</div>
      </div>
      {showContent && (
        <div className={`sh-ribbon-panel${collapsed ? ' peek' : ''}`} onMouseLeave={() => collapsed && !open && setPeek(false)}>
          {layout === 'toolbar' ? (
            <ToolbarRow tab={tab} env={env} />
          ) : (
            <div className="sh-ribbon-content" ref={content} data-testid={`ribbon-${tab.id}`}>
              {groups.map((g, i) => (
                <GroupView key={g.id} g={g} env={env} collapsed={i >= groups.length - collapsedCount} />
              ))}
            </div>
          )}
          {!phone && (
            <button type="button" className="sh-collapse" aria-label={collapsed ? 'Pin the ribbon' : 'Collapse the ribbon'} title={collapsed ? 'Pin the ribbon' : 'Collapse the ribbon'} data-testid="ribbon-collapse" onMouseDown={keepFocus} onClick={() => { setPeek(false); onCollapse(!collapsed); }}>
              <Icon name={collapsed ? 'push_pin' : 'expand_less'} />
            </button>
          )}
        </div>
      )}
    </div>
  );
}

export type { MenuItem };
