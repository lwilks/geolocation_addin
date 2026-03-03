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
                linkType.LoadFrom(targetModelPath, new WorksetConfiguration());
                loadFromSucceeded = true;

                LogHelper.Info($"Relinked type to: {linkInfo.TargetFilePath}");

                // 2. Publish shared coordinates (requires a transaction)
                using (var tx = new Transaction(siteDoc, "Publish Coordinates"))
                {
                    tx.Start();

                    var linkElementId = new LinkElementId(linkInfo.InstanceId);
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
                // 3. Restore original link — use Reload() for cloud links
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
        /// Strategy B: Extract transform from the link instance and apply as project position
        /// on the copied model. Numerically correct but doesn't establish full Revit coordinate linkage.
        /// </summary>
        public static bool PublishViaTransform(Document copiedDoc, Transform linkTransform)
        {
            try
            {
                var origin = linkTransform.Origin;
                var basisX = linkTransform.BasisX;

                // Extract rotation angle from BasisX (angle from east in radians)
                double angle = Math.Atan2(basisX.Y, basisX.X);

                using (var tx = new Transaction(copiedDoc, "Set Shared Coordinates"))
                {
                    tx.Start();

                    var projectLocation = copiedDoc.ActiveProjectLocation;
                    var position = new ProjectPosition(
                        origin.X,   // easting
                        origin.Y,   // northing
                        origin.Z,   // elevation
                        angle        // angle from true north
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
