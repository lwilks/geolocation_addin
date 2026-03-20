using System;
using Autodesk.Revit.DB;
using GeolocationAddin.Helpers;
using GeolocationAddin.Models;

namespace GeolocationAddin.Core
{
    public static class CoordinatePublisher
    {
        /// <summary>
        /// Strategy A: Relink the type to the copied file, publish coordinates, then restore.
        /// This establishes a named shared coordinate position in the copy, enabling
        /// Revit's "By Shared Coordinates" linking between output models.
        /// LoadFrom must be called outside any transaction. PublishCoordinates needs a transaction.
        /// </summary>
        public static bool PublishViaRelink(Document siteDoc, LinkInstanceInfo linkInfo)
        {
            var linkType = linkInfo.LinkType;
            bool loadFromSucceeded = false;

            try
            {
                // 1. Relink to the copied file (must be outside any transaction)
                var targetModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkInfo.TargetFilePath);
                var worksetConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
                linkType.LoadFrom(targetModelPath, worksetConfig);
                loadFromSucceeded = true;

                LogHelper.Info($"Relinked type to: {linkInfo.TargetFilePath}");

                // 2. Get the linked document's active ProjectLocation — PublishCoordinates
                //    needs a LinkElementId that identifies both the link instance and the
                //    target ProjectLocation in the linked model.
                var linkInstance = siteDoc.GetElement(linkInfo.InstanceId) as RevitLinkInstance;
                if (linkInstance == null)
                    throw new InvalidOperationException("Could not resolve RevitLinkInstance after LoadFrom.");

                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null)
                    throw new InvalidOperationException("GetLinkDocument() returned null — linked model not loaded.");

                var linkedLocation = linkDoc.ActiveProjectLocation;
                LogHelper.Info($"Linked doc active location: '{linkedLocation.Name}' (Id={linkedLocation.Id})");
                LogHelper.Info($"Link instance Id={linkInfo.InstanceId}, type Id={linkInfo.TypeId}");

                // 3. Publish shared coordinates (requires a transaction)
                using (var tx = new Transaction(siteDoc, "Publish Coordinates"))
                {
                    tx.Start();

                    // The two-arg constructor LinkElementId(linkInstanceId, linkedElementId) sets
                    // LinkInstanceId and LinkedElementId — both are required by PublishCoordinates.
                    var linkElementId = new LinkElementId(linkInfo.InstanceId, linkedLocation.Id);
                    LogHelper.Info($"LinkElementId: HostElementId={linkElementId.HostElementId}, " +
                                   $"LinkInstanceId={linkElementId.LinkInstanceId}, " +
                                   $"LinkedElementId={linkElementId.LinkedElementId}");

                    siteDoc.PublishCoordinates(linkElementId);

                    tx.Commit();
                }

                LogHelper.Info($"Published coordinates to: {linkInfo.TargetFileName}");

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Strategy A failed for {linkInfo.TargetFileName}: {ex.Message}");
                return false;
            }
            finally
            {
                // 4. Restore original link — use Reload() for cloud links
                if (loadFromSucceeded)
                {
                    try
                    {
                        linkType.Reload();
                        LogHelper.Info("Restored original link via Reload().");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error($"Failed to restore link: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Strategy B: Extract transform from the link instance and combine with the site model's
        /// survey position to compute absolute shared coordinates for the copied model.
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

                using (var tx = new Transaction(copiedDoc, "Set Shared Coordinates"))
                {
                    tx.Start();

                    var projectLocation = copiedDoc.ActiveProjectLocation;
                    var position = new ProjectPosition(
                        absoluteEasting,    // easting in feet
                        absoluteNorthing,   // northing in feet
                        absoluteElevation,  // elevation in feet
                        absoluteAngle       // angle from true north
                    );

                    projectLocation.SetProjectPosition(XYZ.Zero, position);

                    tx.Commit();
                }

                LogHelper.Info("Applied transform-based coordinates (Strategy B).");
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Strategy B failed: {ex.Message}");
                return false;
            }
        }
    }
}
