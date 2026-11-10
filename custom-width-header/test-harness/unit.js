const { resolveHeaderTheme, contrastReport } = require('./theme.cjs');
const D = require('/home/user/Manager/custom-width-header/src/webparts/customWidthHeader/CustomWidthHeaderWebPart.manifest.json').preconfiguredEntries[0].properties;
let fails = 0;
const cases = [];
for (const palette of ['blue', 'indigo', 'green', 'amber']) for (const useGradient of [true, false]) cases.push({ palette, useGradient });
cases.push({ palette: 'custom', backgroundColor: '#ffe066' }, { palette: 'custom', backgroundColor: '#ffffff' }, { palette: 'custom', backgroundColor: '#000000' },
  { palette: 'custom', backgroundColor: '#ffffff', textColor: '#ffffff' }, { palette: 'custom', backgroundColor: '#f4f6fb', textColor: '#0f1c36' });
for (const o of [0, 20, 55, 90]) cases.push({ backgroundType: 'image', overlayOpacity: o });
cases.push({ backgroundType: 'image', overlayColor: '#ffffff', overlayOpacity: 60 }, { backgroundType: 'image', overlayColor: '#ffffff', textColor: '#0f1c36', overlayOpacity: 10 },
  { backgroundType: 'image', imageTextStyle: 'panel', overlayColor: '#ffd400' }, { backgroundType: 'image', imageTextStyle: 'panel', overlayColor: '#0f1c36' });
const titleOnly = { description: '', eyebrow: '', showBackLink: false };
cases.push({ ...titleOnly, palette: 'blue' }, { ...titleOnly, palette: 'amber' }, { ...titleOnly, backgroundType: 'image', overlayOpacity: 0 });
for (const c of cases) {
  const props = { ...D, ...c };
  const t = resolveHeaderTheme(props);
  const r = contrastReport(props);
  const small = props.description.trim() || props.eyebrow.trim() || (props.showBackLink && props.backLinkText.trim());
  const ok = small ? r.worstTitle >= 4.5 && r.worstSmall >= 4.5 : r.worstTitle >= 3;
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'} ${JSON.stringify(c).padEnd(78)} title ${r.worstTitle.toFixed(2)} small ${r.worstSmall.toFixed(2)} | ${t.adjustments.map(a => `${a.label} ${a.from}->${a.to}`).join('; ')}`);
}
process.exit(fails ? 1 : 0);
