/** Message boxes: confirm (OK / Cancel, or Yes / No / Cancel) and a one-line prompt. */
import { useState } from 'react';
import { Icon } from '../Icon';
import { Modal } from '../Modal';

export function ConfirmDialog({ message, title, ok = 'OK', cancel = 'Cancel', alt, danger, onAnswer }: { message: string; title?: string; ok?: string; cancel?: string; alt?: string; danger?: boolean; onAnswer: (v: boolean, alt?: boolean) => void }) {
  return (
    <Modal
      title={title ?? 'Spreadsheet'}
      onClose={() => onAnswer(false)}
      width={420}
      blocking
      testId="confirm-dialog"
      footer={
        <>
          <button type="button" className={`sh-btn-primary${danger ? ' danger' : ''}`} data-testid="confirm-ok" data-autofocus onClick={() => onAnswer(true)}>{ok}</button>
          {alt && <button type="button" className="sh-btn-secondary" data-testid="confirm-alt" onClick={() => onAnswer(false, true)}>{alt}</button>}
          {cancel && <button type="button" className="sh-btn-secondary" data-testid="confirm-cancel" onClick={() => onAnswer(false)}>{cancel}</button>}
        </>
      }
    >
      <div className="sh-message">
        <Icon name={danger ? 'warning' : 'info'} className={danger ? 'warn' : 'info'} />
        <p>{message}</p>
      </div>
    </Modal>
  );
}

export function PromptDialog({ message, value, title, ok = 'OK', onAnswer }: { message: string; value: string; title?: string; ok?: string; onAnswer: (v: string | null) => void }) {
  const [text, setText] = useState(value);
  return (
    <Modal
      title={title ?? message}
      onClose={() => onAnswer(null)}
      width={400}
      blocking
      testId="prompt-dialog"
      footer={
        <>
          <button type="button" className="sh-btn-primary" data-testid="prompt-ok" onClick={() => onAnswer(text)}>{ok}</button>
          <button type="button" className="sh-btn-secondary" onClick={() => onAnswer(null)}>Cancel</button>
        </>
      }
    >
      <div className="sh-form">
        <label className="sh-field">
          <span>{message}</span>
          <input className="sh-input" value={text} data-testid="prompt-input" onFocus={(e) => e.target.select()} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && onAnswer(text)} />
        </label>
      </div>
    </Modal>
  );
}
