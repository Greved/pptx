import type { SlideSvgResult } from '@office-kit/pptx-preview';

const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';

/** Paint renderer output while retaining unchanged root objects and their text layout. */
export function svgPaint(
  node: HTMLDivElement,
  svg: string | SlideSvgResult,
): { update: (svg: string | SlideSvgResult) => void; destroy: () => void } {
  let previous: string | undefined;
  let plan:
    | { source: SlideSvgResult; root: Element; fragments: readonly ChildNode[][] }
    | undefined;
  const template = node.ownerDocument.createElement('template');
  function parseFragment(opening: string, fragment: string): ChildNode[] {
    template.innerHTML = opening + fragment + '</svg>';
    return [...(template.content.firstElementChild?.childNodes ?? [])];
  }
  function paintPlan(source: SlideSvgResult): void {
    const oldPlan = plan;
    let root: Element;
    if (oldPlan && oldPlan.source.opening === source.opening && node.firstChild === oldPlan.root) {
      root = oldPlan.root;
    } else {
      template.innerHTML = source.opening + '</svg>';
      const parsedRoot = template.content.firstElementChild;
      if (!parsedRoot) return;
      root = parsedRoot;
      node.replaceChildren(root);
    }
    const reuse = oldPlan?.root === root ? oldPlan : undefined;
    const fragments: ChildNode[][] = [];
    for (
      let index = 0;
      index < Math.max(source.fragments.length, reuse?.fragments.length ?? 0);
      index++
    ) {
      const fragment = source.fragments[index];
      const oldNodes = reuse?.fragments[index] ?? [];
      if (fragment !== undefined && reuse?.source.fragments[index] === fragment) {
        fragments.push(oldNodes);
        continue;
      }
      const newNodes = fragment === undefined ? [] : parseFragment(source.opening, fragment);
      const anchor =
        reuse?.fragments
          .slice(index)
          .flat()
          .find((child) => child.parentNode === root) ?? null;
      for (const child of newNodes) root.insertBefore(child, anchor);
      for (const child of oldNodes) child.remove();
      if (fragment !== undefined) fragments.push(newNodes);
    }
    plan = { source, root, fragments };
  }
  function update(source: string | SlideSvgResult): void {
    const value = typeof source === 'string' ? source : source.svg;
    if (previous === value) return;
    previous = value;
    if (typeof source !== 'string') {
      paintPlan(source);
      return;
    }
    plan = undefined;
    template.innerHTML = value;
    const oldRoot = node.firstElementChild;
    const newRoot = template.content.firstElementChild;
    // Errors/comments and changed SVG root attributes retain the renderer's exact
    // output. In particular, a new viewBox must never keep an old slide surface.
    if (
      node.childNodes.length !== 1 ||
      template.content.childNodes.length !== 1 ||
      oldRoot?.namespaceURI !== SVG_NAMESPACE ||
      newRoot?.namespaceURI !== SVG_NAMESPACE ||
      oldRoot.localName !== 'svg' ||
      newRoot.localName !== 'svg' ||
      !oldRoot.cloneNode(false).isEqualNode(newRoot.cloneNode(false))
    ) {
      node.replaceChildren(template.content);
      return;
    }
    const oldChildren = [...oldRoot.childNodes];
    const newChildren = [...newRoot.childNodes];
    for (let index = 0; index < Math.max(oldChildren.length, newChildren.length); index++) {
      const oldChild = oldChildren[index];
      const newChild = newChildren[index];
      if (!newChild) {
        oldChild?.remove();
        continue;
      }
      if (!oldChild) {
        oldRoot.append(newChild);
        continue;
      }
      if (!oldChild.isEqualNode(newChild)) oldChild.replaceWith(newChild);
    }
  }
  update(svg);
  return {
    update,
    destroy() {
      previous = undefined;
      plan = undefined;
      node.replaceChildren();
    },
  };
}
