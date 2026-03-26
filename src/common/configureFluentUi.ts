import { setIconOptions } from '@fluentui/react/lib/Styling';

declare global {
  interface Window {
    __opIntranetFluentUiConfigured?: boolean;
  }
}

export const configureFluentUi = (): void => {
  if (window.__opIntranetFluentUiConfigured) {
    return;
  }

  setIconOptions({ disableWarnings: true });
  window.__opIntranetFluentUiConfigured = true;
};
