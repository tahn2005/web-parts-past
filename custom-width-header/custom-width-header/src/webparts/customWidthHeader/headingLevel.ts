/**
 * Picks the title's heading level from the page around the web part, so the
 * page outline never skips a level whatever kind of SharePoint page it is on:
 *
 *  - Article pages show SharePoint's own page title as the page's main (level 1)
 *    heading above the canvas, so the header becomes a level 2 heading under it.
 *  - Home-style pages have no title area, so the header is the level 1 heading.
 */

/** True if assistive technology would announce this element. */
function isExposed(element: Element): boolean {
  if (element.closest('[aria-hidden="true"]')) return false;
  if (element.getClientRects().length === 0) return false; // display:none on it or an ancestor
  return getComputedStyle(element).visibility !== 'hidden';
}

/** Everything announced as a level-1 heading: <h1> elements and ARIA headings marked level 1. */
function levelOneHeadings(): Element[] {
  return Array.from(document.querySelectorAll('h1, [role="heading"][aria-level="1"]'));
}

export function detectHeadingLevel(anchor: HTMLElement): '1' | '2' {
  const hasMainHeadingAbove = levelOneHeadings().some(
    (heading) =>
      !anchor.contains(heading) &&
      // eslint-disable-next-line no-bitwise
      (heading.compareDocumentPosition(anchor) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0 &&
      isExposed(heading)
  );
  return hasMainHeadingAbove ? '2' : '1';
}
