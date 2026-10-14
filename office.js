export let Office, Word;
let ready;

export default function initOffice() {
  return ready ??= (async () => {
    if (!window.Office) await new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
    await window.Office.onReady();
    ({ Office, Word } = window);
  })();
}
