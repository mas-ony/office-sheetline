/** The full-page app: the editor with the options and page connections of the shell. */
import { SheetlineEditor } from '@sheetline/react';
import { useState } from 'react';
import { connectShell, shellOptions } from '@sheetline/web-shared/shell';

export function App() {
  const [options] = useState(shellOptions);
  return <SheetlineEditor {...options} onReady={connectShell} />;
}
