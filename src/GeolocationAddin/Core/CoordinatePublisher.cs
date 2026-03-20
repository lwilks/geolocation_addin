using System;
using System.IO;
using Autodesk.Revit.DB;
using GeolocationAddin.Helpers;
using GeolocationAddin.Models;

namespace GeolocationAddin.Core
{
    public static class CoordinatePublisher
    {
        /// <summary>
        /// Sets shared coordinates on the copied document (must be open as a primary document)
        /// and creates a named ProjectLocation matching the site model to enable
        /// Revit's "By Shared Coordinates" linking between output models.
        /// </summary>
        public static bool PublishViaTransform(Document siteDoc, Document copiedDoc, Transform linkTransform)
        {
            try
            {
                // Get the site model's survey position (where its internal origin sits in survey coords)
                var siteLocation = siteDoc.ActiveProjectLocation;
                var sitePos = siteLocation.GetProjectPosition(XYZ.Zero);

                LogHelper.Info($"Site survey position: E={sitePos.EastWest:F4}, N={sitePos.NorthSouth:F4}, " +
                               $"Elev={sitePos.Elevation:F4}, Angle={sitePos.Angle:F6} rad");

                // Link offset relative to site model's internal origin (in feet)
                var dx = linkTransform.Origin.X;
                var dy = linkTransform.Origin.Y;
                var dz = linkTransform.Origin.Z;

                LogHelper.Info($"Link transform offset: dX={dx:F4}, dY={dy:F4}, dZ={dz:F4}");

                // Rotate the link offset by the site's true north angle to get survey-aligned offset
                double cosA = Math.Cos(sitePos.Angle);
                double sinA = Math.Sin(sitePos.Angle);

                double absoluteEasting = sitePos.EastWest + dx * cosA - dy * sinA;
                double absoluteNorthing = sitePos.NorthSouth + dx * sinA + dy * cosA;
                double absoluteElevation = sitePos.Elevation + dz;

                // Link's own rotation relative to site
                double linkRotation = Math.Atan2(linkTransform.BasisX.Y, linkTransform.BasisX.X);
                double absoluteAngle = sitePos.Angle + linkRotation;

                LogHelper.Info($"Absolute coordinates: E={absoluteEasting:F4}, N={absoluteNorthing:F4}, " +
                               $"Elev={absoluteElevation:F4}, Angle={absoluteAngle:F6} rad");

                var position = new ProjectPosition(
                    absoluteEasting,
                    absoluteNorthing,
                    absoluteElevation,
                    absoluteAngle
                );

                using (var tx = new Transaction(copiedDoc, "Set Shared Coordinates"))
                {
                    tx.Start();

                    // 1. Set coordinates on the active (Internal) location
                    var projectLocation = copiedDoc.ActiveProjectLocation;
                    projectLocation.SetProjectPosition(XYZ.Zero, position);

                    // 2. Create a named ProjectLocation matching the site model.
                    //    This enables "By Shared Coordinates" linking — Revit matches
                    //    position names across models. All copies get the same name
                    //    from the site model, so they can link to each other.
                    try
                    {
                        string siteName = Path.GetFileNameWithoutExtension(siteDoc.Title);
                        var namedLocation = ProjectLocation.Create(copiedDoc, copiedDoc.SiteLocation, siteName);
                        namedLocation.SetProjectPosition(XYZ.Zero, position);
                        LogHelper.Info($"Created named shared position: '{siteName}'");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Info($"Could not create named position: {ex.Message}");
                    }

                    tx.Commit();
                }

                LogHelper.Info("Applied shared coordinates.");
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Coordinate publishing failed: {ex.Message}");
                return false;
            }
        }
    }
}
