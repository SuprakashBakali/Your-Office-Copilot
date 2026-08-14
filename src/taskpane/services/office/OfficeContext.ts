import { OfficeContextInfo, SupportedApiInfo, OfficeHostType } from '../../types';

export function detectOfficeHost(): OfficeHostType {
  if (typeof Office === 'undefined' || !Office.context || !Office.context.host) {
    return 'Unknown';
  }
  const host = Office.context.host;
  switch (host) {
    case Office.HostType.Excel:
      return 'Excel';
    default:
      return 'Unknown';
  }
}

export function detectSupportedApis(): SupportedApiInfo {
  const reqs = typeof Office !== 'undefined' && Office.context ? Office.context.requirements : null;
  
  // Helper to check support safely
  const check = (api: string, version: string) => {
    return reqs ? reqs.isSetSupported(api, version) : false;
  };

  return {
    excelApi1_1: check('ExcelApi', '1.1'),
    excelApi1_2: check('ExcelApi', '1.2'),
    excelApi1_3: check('ExcelApi', '1.3'),
    excelApi1_4: check('ExcelApi', '1.4'),
    excelApi1_5: check('ExcelApi', '1.5'),
    excelApi1_6: check('ExcelApi', '1.6'),
    excelApi1_7: check('ExcelApi', '1.7'),
    excelApi1_8: check('ExcelApi', '1.8'),
    excelApi1_9: check('ExcelApi', '1.9'),
    excelApi1_10: check('ExcelApi', '1.10'),
    excelApi1_11: check('ExcelApi', '1.11'),
    excelApi1_12: check('ExcelApi', '1.12'),
    excelApi1_13: check('ExcelApi', '1.13'),
    excelApi1_14: check('ExcelApi', '1.14'),
    excelApi1_14: check('ExcelApi', '1.14'),
  };
}

export function getOfficeContext(): OfficeContextInfo {
  const isReady = typeof Office !== 'undefined' && Office.context != null;
  const platform = isReady ? Office.context.platform : 'Unknown';
  const version = (isReady && Office.context.diagnostics) ? Office.context.diagnostics.version : 'Unknown';
  
  return {
    host: detectOfficeHost(),
    platform: platform as any, // casting based on type assumptions
    version: version,
    isReady: isReady,
    supportedApis: detectSupportedApis()
  };
}

export function getOfficeVersionName(context: OfficeContextInfo): string {
  if (!context.isReady) return 'Unknown Version';

  const platform = context.platform;

  // Compare against the enum directly — String() on a numeric enum is
  // fragile and version-dependent. Office.context.platform returns the
  // enum value, but it's typed as `any` so we cast for the comparison.
  if (platform as any === Office.PlatformType.OfficeOnline) {
    return 'Office Online';
  }

  if (platform as any === Office.PlatformType.Mac || platform as any === Office.PlatformType.PC) {
    if (context.host === 'Excel') {
      if (context.supportedApis.excelApi1_14) return 'Microsoft 365';
      if (context.supportedApis.excelApi1_13) return 'Office 2021';
      if (context.supportedApis.excelApi1_9) return 'Office 2019';
      if (context.supportedApis.excelApi1_3) return 'Office 2016';
    }
  }

  return `Office (${platform})`;
}
