import '@primer/primitives/dist/css/primitives.css';
import '@primer/primitives/dist/css/functional/themes/light.css';
import { BaseStyles, PageLayout, ThemeProvider } from '@primer/react';
import type { RefObject } from 'react';
import { createRoot } from 'react-dom/client';
import { flushSync } from 'react-dom';
import { Toolbar } from './Toolbar.tsx';
import type { ToolbarActions, ToolbarState } from './Toolbar.tsx';

/**
 * Renders the page layout synchronously, so `#grid-container`, `#add-sheet-button` and `#sheet-tabs` are in the DOM
 * when this function returns. Their contents are owned by the caller, React never re-renders them.
 *
 * @param returnFocusRef Element to focus after a toolbar menu closes.
 */
export function mountUI(
  container: HTMLElement,
  actions: ToolbarActions,
  initialState: ToolbarState,
  returnFocusRef: RefObject<HTMLElement>
) {
  let state = initialState;
  const listeners = new Set<() => void>();
  const store = {
    getState: () => state,
    subscribe(listener: () => void) {
      listeners.add(listener);

      return () => listeners.delete(listener);
    }
  };

  flushSync(() =>
    createRoot(container).render(
      <ThemeProvider colorMode="day">
        <BaseStyles className="app">
          <PageLayout className="app__layout" columnGap="none" containerWidth="full" padding="none" rowGap="none">
            <PageLayout.Header className="app__header" divider="line" padding="condensed">
              <Toolbar actions={actions} returnFocusRef={returnFocusRef} store={store} />
            </PageLayout.Header>
            <PageLayout.Content className="app__content" padding="none" width="full">
              <div className="grid-container" id="grid-container" />
            </PageLayout.Content>
            <PageLayout.Footer className="app__footer" divider="line" padding="none">
              <div className="sheet-bar">
                <button
                  aria-label="Add sheet"
                  className="sheet-bar__add"
                  id="add-sheet-button"
                  title="Add sheet"
                  type="button"
                >
                  +
                </button>
                <div aria-label="Sheets" className="sheet-bar__tabs" id="sheet-tabs" role="tablist" />
              </div>
            </PageLayout.Footer>
          </PageLayout>
        </BaseStyles>
      </ThemeProvider>
    )
  );

  return {
    update(nextState: Partial<ToolbarState>) {
      if (Object.entries(nextState).some(([key, value]) => !Object.is(state[key as keyof ToolbarState], value))) {
        state = { ...state, ...nextState };
        listeners.forEach(listener => listener());
      }
    }
  };
}
