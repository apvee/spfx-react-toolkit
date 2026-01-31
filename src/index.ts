/**
 * SPFx React Toolkit - React hooks and providers for SharePoint Framework
 *
 * A comprehensive library providing React hooks and context providers
 * for SPFx (SharePoint Framework) development. Simplifies access to
 * SharePoint context, services, and APIs with type-safe React patterns.
 *
 * @packageDocumentation
 *
 * @example Basic usage with WebPart
 * ```tsx
 * import { SPFxWebPartProvider, useSPFxPageContext } from 'spfx-react-toolkit';
 *
 * function MyComponent() {
 *   const pageContext = useSPFxPageContext();
 *   return <div>Site: {pageContext.web.title}</div>;
 * }
 *
 * // In WebPart render():
 * const element = React.createElement(
 *   SPFxWebPartProvider,
 *   { instance: this },
 *   React.createElement(MyComponent)
 * );
 * ReactDom.render(element, this.domElement);
 * ```
 */
export * from './core';
export * from './hooks';