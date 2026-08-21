using System;
using System.Globalization;

namespace MSOfficeAIAssistant.Core
{
    public enum OfficeVersion
    {
        Unknown = 0,
        Office2010 = 14,
        Office2013 = 15,
        Office2016 = 16,
        Office2019 = 17,
        Office2021 = 18,
        Office365 = 365
    }

    public enum OfficeFeature
    {
        CustomTaskPanes,
        RibbonExtensibility,
        TrackChanges,
        ModernContentControls,
        CoAuthoring,
        SvgIcons
    }

    public static class VersionDetector
    {
        public static OfficeVersion DetectVersion(object app)
        {
            if (app == null) return OfficeVersion.Unknown;

            try
            {
                dynamic dynamicApp = app;
                string versionStr = Convert.ToString(dynamicApp.Version, CultureInfo.InvariantCulture);
                if (string.IsNullOrEmpty(versionStr))
                {
                    return OfficeVersion.Unknown;
                }

                string[] parts = versionStr.Split('.');
                int major;
                if (int.TryParse(parts[0], out major))
                {
                    switch (major)
                    {
                        case 14:
                            return OfficeVersion.Office2010;
                        case 15:
                            return OfficeVersion.Office2013;
                        case 16:
                            try
                            {
                                string buildStr = Convert.ToString(dynamicApp.Build, CultureInfo.InvariantCulture);
                                if (!string.IsNullOrEmpty(buildStr))
                                {
                                    string[] buildParts = buildStr.Split('.');
                                    // Office version discrimination via build number (16.0.xxxx format):
                                    // 16.0.2xxxx (20000+) = Office 2021 (perpetual license)
                                    // 16.0.1xxxx (10000+) = Office 365 / Office 2019 (subscription or recent perpetual)
                                    // 16.0.xxxx  (<10000) = Office 2016
                                    if (buildParts.Length > 1)
                                    {
                                        int buildNumber;
                                        if (int.TryParse(buildParts[1], out buildNumber))
                                        {
                                            if (buildNumber >= 20000) return OfficeVersion.Office2021;
                                            if (buildNumber >= 10000) return OfficeVersion.Office365;
                                        }
                                    }
                                }
                            }
                            catch { }
                            return OfficeVersion.Office2016;
                        default:
                            if (major > 16) return OfficeVersion.Office365;
                            return OfficeVersion.Unknown;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to detect Office version: {0}", ex.Message));
            }

            return OfficeVersion.Unknown;
        }

        public static bool SupportsFeature(OfficeFeature feature, OfficeVersion version)
        {
            switch (feature)
            {
                case OfficeFeature.CustomTaskPanes:
                case OfficeFeature.RibbonExtensibility:
                case OfficeFeature.TrackChanges:
                    return version >= OfficeVersion.Office2010;

                case OfficeFeature.ModernContentControls:
                    return version >= OfficeVersion.Office2013;

                case OfficeFeature.SvgIcons:
                case OfficeFeature.CoAuthoring:
                    return version >= OfficeVersion.Office2019 || version == OfficeVersion.Office365;

                default:
                    return true;
            }
        }
    }
}
