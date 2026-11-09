<script setup lang="ts">
/** Insert Hyperlink: a web page, a place in this document or an e-mail address. */
import { cmd } from '@sheetline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../../context';
import Dialog from '../Dialog.vue';
import Field from '../Field.vue';
import ListBox from '../ListBox.vue';
import Tabs from '../Tabs.vue';

type LinkTab = 'web' | 'doc' | 'email';

const emit = defineEmits<{ close: [] }>();

const TABS: { id: LinkTab; label: string }[] = [
  { id: 'web', label: 'Web Page' },
  { id: 'doc', label: 'This Document' },
  { id: 'email', label: 'Email Address' },
];

const ui = useUI();
const { doc } = ui.value;
const { r, c } = doc.selection.active;
const link = doc.sheet.hyperlinkAt(r, c);
const tab = ref<LinkTab>(link?.location ? 'doc' : link?.target?.startsWith('mailto:') ? 'email' : 'web');
const text = ref(doc.displayText(r, c) || link?.display || '');
const tip = ref(link?.tooltip ?? '');
const url = ref(link?.target && !link.target.startsWith('mailto:') ? link.target : '');
const sheetName = ref(link?.location?.includes('!') ? link.location.split('!')[0].replace(/^'|'$/g, '') : doc.sheet.name);
const cellRef = ref(link?.location?.includes('!') ? link.location.split('!')[1] : 'A1');
const email = ref(link?.target?.startsWith('mailto:') ? decodeURIComponent(link.target.slice(7).split('?')[0]) : '');
const subject = ref(link?.target?.includes('subject=') ? decodeURIComponent(link.target.split('subject=')[1]) : '');
const sheets = computed(() => ui.value.doc.wb.sheets.filter((s) => s.state === 'visible').map((s) => s.name));
const same = (x: string) => x;
const val = (e: Event) => (e.target as HTMLInputElement).value;

const removeLink = () => {
  ui.value.doc.removeLink();
  emit('close');
};
const ok = () => {
  const ctx = ui.value;
  const { doc } = ctx;
  let target: string;
  if (tab.value === 'web') {
    const u = url.value.trim();
    if (!u) {
      ctx.notify('Enter an address.');
      return false;
    }
    target = /^[a-z][a-z0-9+.-]*:/i.test(u) ? u : `http://${u}`;
  } else if (tab.value === 'doc') {
    const name = sheetName.value;
    const q = /[^\w]/.test(name) ? `'${name.replace(/'/g, "''")}'` : name;
    target = `#${q}!${cellRef.value.trim() || 'A1'}`;
  } else {
    if (!email.value.trim()) {
      ctx.notify('Enter an e-mail address.');
      return false;
    }
    target = `mailto:${email.value.trim()}${subject.value ? `?subject=${encodeURIComponent(subject.value)}` : ''}`;
  }
  cmd.commit(ctx);
  const current = doc.displayText(r, c);
  const t = text.value;
  doc.setLink(target, t && t !== current ? t : current ? undefined : t || target.replace(/^#/, ''), r, c, doc.sheet, tip.value || undefined);
  return true;
};
</script>

<template>
  <Dialog title="Insert Hyperlink" :width="480" test-id="hyperlink-dialog" :on-ok="ok" @close="emit('close')">
    <template v-if="link" #footer-left>
      <button type="button" class="sh-btn" data-testid="hyperlink-remove" @click="removeLink">Remove</button>
    </template>
    <Field label="Text to display:" wide>
      <input :value="text" data-testid="hyperlink-text" @input="text = val($event)" />
    </Field>
    <Field label="ScreenTip:" wide>
      <input :value="tip" @input="tip = val($event)" />
    </Field>
    <Tabs :value="tab" :tabs="TABS" test-id="hyperlink-tabs" @change="tab = $event" />
    <div class="sh-tab-body">
      <Field v-if="tab === 'web'" label="Address:" wide>
        <input :value="url" placeholder="URL, like http://www.example.com" inputmode="url" data-testid="hyperlink-url" @input="url = val($event)" />
      </Field>
      <div v-if="tab === 'doc'" class="sh-stack">
        <Field label="Type the cell reference:" wide>
          <input :value="cellRef" data-testid="hyperlink-ref" @input="cellRef = val($event)" />
        </Field>
        <div class="sh-label">Or select a place in this document:</div>
        <ListBox v-slot="{ item }" :items="sheets" :value="sheetName" :key-of="same" :height="120" label="Sheets" @change="sheetName = $event">{{ item }}</ListBox>
      </div>
      <div v-if="tab === 'email'" class="sh-stack">
        <Field label="E-mail address:" wide>
          <input :value="email" inputmode="email" data-testid="hyperlink-email" @input="email = val($event)" />
        </Field>
        <Field label="Subject:" wide>
          <input :value="subject" @input="subject = val($event)" />
        </Field>
      </div>
    </div>
  </Dialog>
</template>
