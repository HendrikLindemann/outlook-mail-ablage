// Erzeugt die Symbole (16–128 px) als PNG. Nutzt den installierten Microsoft Edge.
const { chromium } = require("playwright");
(async () => {
  const b = await chromium.launch({ channel: "msedge" });
  for (const size of [16, 32, 64, 80, 128]) {
    const p = await b.newPage({ viewport: { width: size, height: size } });
    await p.setContent(`<body style="margin:0;background:transparent"><svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 32 32">
      <path d="M2 8a3 3 0 0 1 3-3h7l3 3h12a3 3 0 0 1 3 3v14a3 3 0 0 1-3 3H5a3 3 0 0 1-3-3z" fill="#0f6cbd"/>
      <rect x="8" y="13" width="16" height="10" rx="1.5" fill="#fff"/>
      <path d="M8.6 13.8 16 19l7.4-5.2" fill="none" stroke="#0f6cbd" stroke-width="1.8" stroke-linejoin="round"/></svg></body>`);
    await p.screenshot({ path: `public/assets/icon-${size}.png`, omitBackground: true });
  }
  await b.close();
})();
