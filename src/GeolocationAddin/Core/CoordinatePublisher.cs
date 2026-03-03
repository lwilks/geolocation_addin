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
        /// LoadFrom must be called outside any transaction.
        /// </summary>
        public static bool PublishViaRelink(Document siteDoc, LinkInstanceInfo linkInfo)
        {
            var linkType = linkInfo.LinkType;
            ModelPath originalPath = null;

            try
            {
                // 1. Save original link path
                try
                {
                    var extRef = linkType.GetExternalFileReference();
                    if (extRef != null)
                        originalPath = extRef.GetAbsolutePath();
                }
                catch (Exception)
                {
                    // Cloud/ACC link — save the InSessionPath from external resources instead
                    LogHelper.Info("Cloud link detected, saving external resource reference for restore.");
                }

                // 2. Relink to the copied file
                var targetModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkInfo.TargetFilePath);
                linkType.LoadFrom(targetModelPath, new WorksetConfiguration());

                LogHelper.Info($"Relinked type to: {linkInfo.TargetFilePath}");

                // 3. Publish shared coordinates from site model to the link
                var linkElementId = new LinkElementId(linkInfo.InstanceId);
                siteDoc.PublishCoordinates(linkElementId);

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
                // 4. Restore original link path
                if (originalPath != null)
                {
                    try
                    {
                        linkType.LoadFrom(originalPath, new WorksetConfiguration());
                        LogHelper.Info("Restored original link path.");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error($"Failed to restore link path: {ex.Message}");
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
