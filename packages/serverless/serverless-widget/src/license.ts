let globalLicenseKey: string | null = null;

export function setLicenseKey(licenseKey: string) {
  globalLicenseKey = licenseKey;
}

export function getLicenseKey(): string | null {
  return globalLicenseKey;
}
