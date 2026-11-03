// Automated Test Suite for Gold Standard Responsive Vertical Scrolling & Header
const fs = require('fs');
const path = require('path');

console.log('=== TEST SUITE: Vertical Scrolling & Responsive Header Compliance ===');

// 1. Audit styles.css
const css = fs.readFileSync(path.join(__dirname, '../mobile-web/styles.css'), 'utf8');

console.log('\n--- 1. Viewport & Body Scrolling Standards ---');
const hasSmoothScroll = css.includes('scroll-behavior: smooth;');
const hasDvh = css.includes('100dvh');
const hasOverscrollContain = css.includes('overscroll-behavior-y: contain;');
const hasTouchScrolling = css.includes('-webkit-overflow-scrolling: touch;');

console.log('  Smooth scrolling enabled:', hasSmoothScroll);
console.log('  Dynamic Viewport Height (100dvh):', hasDvh);
console.log('  Overscroll Behavior Contain:', hasOverscrollContain);
console.log('  Touch Momentum Scrolling:', hasTouchScrolling);

if (!hasSmoothScroll || !hasDvh || !hasOverscrollContain || !hasTouchScrolling) {
  throw new Error('Missing core viewport/body scrolling rules');
}

console.log('\n--- 2. State-Aware Compact Sticky Header Rules ---');
const hasIsCompact = css.includes('.gfmp-header.is-compact');
const hasLogoShrink = css.includes('.gfmp-header.is-compact .brand-logo-img');
const hasSubTitleHide = css.includes('.gfmp-header.is-compact .brand-text .sub-title');
const hasBtnRefreshCompact = css.includes('.gfmp-header.is-compact .btn-refresh-rates');
const hasScrollTimelineSupports = css.includes('animation-timeline: scroll()');

console.log('  .is-compact class defined:', hasIsCompact);
console.log('  Logo shrinks on scroll:', hasLogoShrink);
console.log('  Subtitle hidden on compact scroll:', hasSubTitleHide);
console.log('  Refresh button smoothly compacts on scroll:', hasBtnRefreshCompact);
console.log('  Native CSS scroll-timeline @supports:', hasScrollTimelineSupports);

if (!hasIsCompact || !hasLogoShrink || !hasSubTitleHide || !hasBtnRefreshCompact || !hasScrollTimelineSupports) {
  throw new Error('Missing compact sticky header rules in styles.css');
}

console.log('\n--- 3. Table Responsive Horizontal Scroll Isolation ---');
const hasTableCardScroll = css.includes('.table-card') && css.includes('overflow-x: auto;');
const hasTableCardOverscroll = css.includes('overscroll-behavior-x: contain;');
console.log('  Table card horizontal scroll isolation:', hasTableCardScroll);
console.log('  Table card horizontal overscroll isolation:', hasTableCardOverscroll);

if (!hasTableCardScroll || !hasTableCardOverscroll) {
  throw new Error('Table card missing horizontal scroll isolation');
}

console.log('\n--- 4. Comprehensive Responsive Media Queries ---');
const hasMobilePortraitMQ = css.includes('@media (max-width: 480px)');
const hasSmallMobileMQ = css.includes('@media (max-width: 360px)');
const hasLandscapeMQ = css.includes('@media (orientation: landscape) and (max-height: 520px)');
const hasReducedMotionMQ = css.includes('@media (prefers-reduced-motion: reduce)');

console.log('  Mobile Portrait (<= 480px):', hasMobilePortraitMQ);
console.log('  Compact Mobile (<= 360px):', hasSmallMobileMQ);
console.log('  Mobile Landscape (<= 520px):', hasLandscapeMQ);
console.log('  Prefers Reduced Motion:', hasReducedMotionMQ);

if (!hasMobilePortraitMQ || !hasSmallMobileMQ || !hasLandscapeMQ || !hasReducedMotionMQ) {
  throw new Error('Missing responsive media queries');
}

console.log('\n--- 5. JavaScript Scroll Controller in app.js ---');
const appJs = fs.readFileSync(path.join(__dirname, '../mobile-web/app.js'), 'utf8');
const hasInitScroll = appJs.includes('initScrollResponsiveness');
const hasTickingRaf = appJs.includes('requestAnimationFrame');
const hasPassiveListener = appJs.includes('{ passive: true }');

console.log('  initScrollResponsiveness present:', hasInitScroll);
console.log('  requestAnimationFrame throttling:', hasTickingRaf);
console.log('  passive scroll listener:', hasPassiveListener);

if (!hasInitScroll || !hasTickingRaf || !hasPassiveListener) {
  throw new Error('Missing or suboptimal JavaScript scroll controller in app.js');
}

console.log('\n>>> ALL 5 VERTICAL SCROLLING & RESPONSIVE HEADER TESTS PASSED (0 ERRORS) <<<');
