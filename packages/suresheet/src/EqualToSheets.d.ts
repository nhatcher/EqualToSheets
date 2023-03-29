declare namespace EqualToSheets {
  export declare function setLicenseKey(licenseKey: string): void;
  export declare function load(
    workbookId: string,
    element: HTMLElement,
    options?: {
      onLoad?: (workbookId: string, workbookJson: string) => void;
      saveWorkbookPreHook?: (workbookId: string, workbookJson: string) => boolean;
    },
  ): void;
}
