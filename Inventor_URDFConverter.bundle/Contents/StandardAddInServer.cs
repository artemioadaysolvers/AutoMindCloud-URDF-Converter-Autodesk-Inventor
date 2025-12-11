using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;               // Para PNGs
using System.Drawing.Imaging;       // Para guardar PNG
using System.IO;
using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using Inventor;

namespace URDFConverterAddIn
{
    // GUID NUEVO: d507347f-7ff3-4dc0-a66d-5e596f9e9d89
    [Guid("d507347f-7ff3-4dc0-a66d-5e596f9e9d89")]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application _invApp;

        // Dos botones: VeryLowOptimized y DisplayMesh
        private ButtonDefinition _exportUrdfVlqButton;
        private ButtonDefinition _exportUrdfDisplayButton;

        // ClientId del AddIn: DEBE ser el mismo GUID que arriba, pero con llaves
        private const string AddInClientId = "{d507347f-7ff3-4dc0-a66d-5e596f9e9d89}";

        // ----------------------------------------------------
        //  Activate: se ejecuta cuando Inventor carga el AddIn
        // ----------------------------------------------------
        public void Activate(ApplicationAddInSite AddInSiteObject, bool FirstTime)
        {
            _invApp = AddInSiteObject.Application;

            Debug.WriteLine("[URDF][SYS] Activate() llamado. FirstTime = " + FirstTime);

            try
            {
                CommandManager cmdMgr = _invApp.CommandManager;
                ControlDefinitions controlDefs = cmdMgr.ControlDefinitions;

                // ---------------------------------------------
                // 1) DEFINICIONES DE BOTONES (ButtonDefinition)
                // ---------------------------------------------

                // Botón 1: Very Low Quality Optimized (VLQ)
                _exportUrdfVlqButton = null;
                try
                {
                    _exportUrdfVlqButton =
                        controlDefs["urdf_export_vlq_cmd"] as ButtonDefinition;
                }
                catch (Exception exLookup)
                {
                    Debug.WriteLine("[URDF][SYS] lookup 'urdf_export_vlq_cmd' lanzó: " + exLookup.Message);
                    _exportUrdfVlqButton = null;
                }

                if (_exportUrdfVlqButton == null)
                {
                    _exportUrdfVlqButton = controlDefs.AddButtonDefinition(
                        "Export URDF (VLQ)",                 // DisplayName
                        "urdf_export_vlq_cmd",               // InternalName (único)
                        CommandTypesEnum.kNonShapeEditCmdType,
                        AddInClientId,                       // 👉 ClientId = GUID del AddIn con llaves
                        "Export URDF with VeryLowOptimized mesh",
                        "Export URDF (VLQ)");
                }

                _exportUrdfVlqButton.OnExecute +=
                    new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfVlqButtonPressed);

                // Botón 2: DisplayMesh (alta calidad)
                _exportUrdfDisplayButton = null;
                try
                {
                    _exportUrdfDisplayButton =
                        controlDefs["urdf_export_display_cmd"] as ButtonDefinition;
                }
                catch (Exception exLookup)
                {
                    Debug.WriteLine("[URDF][SYS] lookup 'urdf_export_display_cmd' lanzó: " + exLookup.Message);
                    _exportUrdfDisplayButton = null;
                }

                if (_exportUrdfDisplayButton == null)
                {
                    _exportUrdfDisplayButton = controlDefs.AddButtonDefinition(
                        "Export URDF (Display)",             // DisplayName
                        "urdf_export_display_cmd",           // InternalName (único)
                        CommandTypesEnum.kNonShapeEditCmdType,
                        AddInClientId,                       // 👉 MISMO ClientId
                        "Export URDF with DisplayMesh-quality mesh",
                        "Export URDF (Display)");
                }

                Debug.WriteLine("[URDF][SYS] ButtonDefinitions creados/obtenidos correctamente.");

                _exportUrdfDisplayButton.OnExecute +=
                    new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed);

                // -------------------------------------------------
                // 2) Añadir los botones a los Ribbons de Part y Assembly
                // -------------------------------------------------
                UserInterfaceManager uiMgr = _invApp.UserInterfaceManager;

                Debug.WriteLine("[URDF][SYS] Creando panels en ribbons...");

                // 2.1) Ribbon de PIEZAS (Part)
                try
                {
                    Ribbon partRibbon      = uiMgr.Ribbons["Part"];
                    RibbonTab toolsTabPart = partRibbon.RibbonTabs["id_TabTools"];

                    RibbonPanel urdfPanelPart = null;
                    try
                    {
                        urdfPanelPart = toolsTabPart.RibbonPanels["urdf_export_panel_part"];
                    }
                    catch
                    {
                        urdfPanelPart = null;
                    }

                    if (urdfPanelPart == null)
                    {
                        urdfPanelPart = toolsTabPart.RibbonPanels.Add(
                            "URDF Export",
                            "urdf_export_panel_part",
                            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                            "",
                            false);
                    }

                    // Añadimos los botones al panel
                    urdfPanelPart.CommandControls.AddButton(_exportUrdfVlqButton);
                    urdfPanelPart.CommandControls.AddButton(_exportUrdfDisplayButton);

                    Debug.WriteLine("[URDF][UI] Panel URDF en Part creado/actualizado correctamente.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[URDF][UI] Error creando panel en Part: " + ex.Message);
                }

                // 2.2) Ribbon de ENSAMBLAJES (Assembly)
                try
                {
                    Ribbon asmRibbon      = uiMgr.Ribbons["Assembly"];
                    RibbonTab toolsTabAsm = asmRibbon.RibbonTabs["id_TabTools"];

                    RibbonPanel urdfPanelAsm = null;
                    try
                    {
                        urdfPanelAsm = toolsTabAsm.RibbonPanels["urdf_export_panel_asm"];
                    }
                    catch
                    {
                        urdfPanelAsm = null;
                    }

                    if (urdfPanelAsm == null)
                    {
                        urdfPanelAsm = toolsTabAsm.RibbonPanels.Add(
                            "URDF Export",
                            "urdf_export_panel_asm",
                            "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
                            "",
                            false);
                    }

                    urdfPanelAsm.CommandControls.AddButton(_exportUrdfVlqButton);
                    urdfPanelAsm.CommandControls.AddButton(_exportUrdfDisplayButton);

                    Debug.WriteLine("[URDF][UI] Panel URDF en Assembly creado/actualizado correctamente.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[URDF][UI] Error creando panel en Assembly: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[URDF][ERR] Activate() EXCEPTION: " + ex.ToString());

                MessageBox.Show(
                    "Error activating URDFConverter AddIn:\n" + ex.Message,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------
        //  Botón VLQ → VeryLowOptimized
        // ----------------------------------------------------
        private void OnExportUrdfVlqButtonPressed(NameValueMap Context)
        {
            try
            {
                UrdfExporter.SetMeshQualityVeryLow();
                UrdfExporter.ExportActiveDocument(_invApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error exporting URDF (VeryLowOptimized):\n" + ex.Message,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------
        //  Botón Display → DisplayMesh (alta calidad)
        // ----------------------------------------------------
        private void OnExportUrdfDisplayButtonPressed(NameValueMap Context)
        {
            try
            {
                UrdfExporter.SetMeshQualityDisplay();
                UrdfExporter.ExportActiveDocument(_invApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error exporting URDF (DisplayMesh):\n" + ex.Message,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void Deactivate()
        {
            try
            {
                if (_exportUrdfVlqButton != null)
                {
                    Marshal.FinalReleaseComObject(_exportUrdfVlqButton);
                    _exportUrdfVlqButton = null;
                }

                if (_exportUrdfDisplayButton != null)
                {
                    Marshal.FinalReleaseComObject(_exportUrdfDisplayButton);
                    _exportUrdfDisplayButton = null;
                }
            }
            catch
            {
            }

            if (_invApp != null)
            {
                try { Marshal.ReleaseComObject(_invApp); }
                catch { }
                _invApp = null;
            }
        }

        public void ExecuteCommand(int CommandID) { }
        public object Automation { get { return null; } }
    }

    // ========================================================================
    //  URDF EXPORTER (parte 1: configuración + entry point)
    // ========================================================================
    public static class UrdfExporter
    {
        // "low"  → VeryLowOptimized  → PNG sólido por .dae
        // "high" → DisplayMesh       → Atlas por .dae (per-face)
        private static string _meshQualityMode = "low";

        public static void SetMeshQualityVeryLow() { _meshQualityMode = "low"; }
        public static void SetMeshQualityDisplay() { _meshQualityMode = "high"; }
        public static string GetMeshQualityMode() { return _meshQualityMode; }

        // Debug flags
        private static bool _DEBUG_SYS        = true;
        private static bool _DEBUG_TRANSFORMS = true;
        private static bool _DEBUG_MESH_TREE  = true;
        private static bool _DEBUG_LINK_JOINT = true;

        private static void DebugLog(string category, string message)
        {
            if (category == "SYS"  && !_DEBUG_SYS)        return;
            if (category == "TFM"  && !_DEBUG_TRANSFORMS) return;
            if (category == "MESH" && !_DEBUG_MESH_TREE)  return;
            if (category == "LINK" && !_DEBUG_LINK_JOINT) return;

            string full = "[URDF][" + category + "] " + message;
            Debug.WriteLine(full);
            Trace.WriteLine(full);
        }

        // =====================================================
        //  ExportActiveDocument
        // =====================================================
        public static void ExportActiveDocument(Inventor.Application invApp)
        {
            if (invApp == null)
            {
                MessageBox.Show("Inventor.Application es nulo.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            Document doc = invApp.ActiveDocument as Document;
            if (doc == null)
            {
                MessageBox.Show("No hay documento activo para exportar.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject &&
                doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
            {
                MessageBox.Show("Solo se soportan documentos de pieza (.ipt) y ensamblaje (.iam).",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string fullPath = string.Empty;
            try { fullPath = doc.FullFileName; }
            catch { fullPath = string.Empty; }

            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show("El documento no tiene ruta de fichero. Guárdalo antes de exportar.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string baseDir  = IOPath.GetDirectoryName(fullPath);
            string baseName = IOPath.GetFileNameWithoutExtension(fullPath);

            DebugLog("SYS",
                "ExportActiveDocument: doc='" + doc.DisplayName +
                "', type=" + doc.DocumentType.ToString() +
                "', path='" + fullPath +
                "', meshMode=" + _meshQualityMode);

            string exportDir = IOPath.Combine(baseDir, "URDF_Export");
            if (!EnsureDirectory(exportDir))
            {
                MessageBox.Show("No se pudo crear la carpeta de exportación:\n" + exportDir,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            string urdfPath = IOPath.Combine(exportDir, baseName + ".urdf");
            DebugLog("SYS", "ExportActiveDocument: exportDir='" + exportDir +
                            "', urdfPath='" + urdfPath + "'");

            try
            {
                // 1) Construir el modelo URDF (links + joints base)
                RobotModel robot = BuildRobotFromDocument(doc, baseName);

                // 2) Exportar geometría + PNG/Atlas por .dae
                ExportGeometryAndTextures(invApp, doc, robot, exportDir);

                // 3) Escribir .urdf
                WriteUrdfFile(robot, urdfPath);

                DebugLog("SYS", "URDF escrito en: " + urdfPath);

                MessageBox.Show("Exportación URDF completada:\n" + urdfPath,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error durante la exportación URDF:\n" + ex.Message,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static bool EnsureDirectory(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            try
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // =====================================================
        //  BuildRobotFromDocument
        // =====================================================
        private static RobotModel BuildRobotFromDocument(Document doc, string baseName)
        {
            RobotModel robot = new RobotModel();
            robot.Name = baseName;

            // base_link
            UrdfLink baseLink = new UrdfLink();
            baseLink.Name = "base_link";
            baseLink.OriginXYZ = new double[] { 0, 0, 0 };
            baseLink.OriginRPY = new double[] { 0, 0, 0 };
            robot.Links.Add(baseLink);

            DebugLog("SYS", "BuildRobotFromDocument: type=" + doc.DocumentType.ToString());

            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                AddPartBodiesAsLinks((PartDocument)doc, robot, baseName);
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AddAssemblyOccurrencesAndBodiesAsLinks((AssemblyDocument)doc, robot);
            }

            DebugLog("SYS",
                "Robot construido: links=" + robot.Links.Count +
                ", joints=" + robot.Joints.Count);

            return robot;
        }











        // =====================================================
        //  PART: un link por SurfaceBody (root_body_i_...)
        // =====================================================
        private static void AddPartBodiesAsLinks(
            PartDocument partDoc,
            RobotModel robot,
            string baseName)
        {
            PartComponentDefinition partDef = partDoc.ComponentDefinition;

            List<SurfaceBody> bodies = new List<SurfaceBody>();
            CollectSurfaceBodiesFromPartDefinition(partDef, bodies);

            DebugLog("MESH",
                "AddPartBodiesAsLinks: part='" + baseName +
                "', SurfaceBodies=" + bodies.Count);

            if (bodies.Count == 0)
            {
                string linkName = "link_" + MakeSafeName(baseName);

                UrdfLink link = new UrdfLink();
                link.Name = linkName;
                link.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                link.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                robot.Links.Add(link);

                UrdfJoint joint = new UrdfJoint();
                joint.Name = "root_" + linkName;
                joint.Type = "fixed";
                joint.ParentLink = "base_link";
                joint.ChildLink = linkName;
                joint.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                joint.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                robot.Joints.Add(joint);

                DebugLog("LINK",
                    "Part sin SurfaceBodies: creado link único '" + linkName + "'");
                return;
            }

            for (int i = 0; i < bodies.Count; i++)
            {
                SurfaceBody b = bodies[i];
                string bodyName = "(null)";
                try
                {
                    if (b != null && !string.IsNullOrEmpty(b.Name))
                        bodyName = b.Name;
                }
                catch { }

                string linkName = "root_body_" +
                                  i.ToString(CultureInfo.InvariantCulture) + "_" +
                                  MakeSafeName(bodyName);

                UrdfLink link2 = new UrdfLink();
                link2.Name = linkName;
                link2.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                link2.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                robot.Links.Add(link2);

                UrdfJoint joint2 = new UrdfJoint();
                joint2.Name = "root_" + linkName;
                joint2.Type = "fixed"; // Para .ipt lo dejamos fijo al base_link
                joint2.ParentLink = "base_link";
                joint2.ChildLink = linkName;
                joint2.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                joint2.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                robot.Joints.Add(joint2);

                DebugLog("LINK",
                    "AddPartBodiesAsLinks: creado link '" + linkName +
                    "' para SurfaceBody[" + i.ToString(CultureInfo.InvariantCulture) + "]");
            }
        }

        // =====================================================
        //  ASSEMBLY: un link por body y por occurrence hoja
        //           nombres únicos: link_<occIndex>_<occName>[_bN]
        //
        //  NOTA: aquí, de momento, todos los joints se marcan como
        //  "fixed". Cuando conectemos propiedades fiables del .iam
        //  (constraints, etc.), este es el sitio para cambiar:
        //  joint.Type → "revolute", "prismatic", "continuous", etc.
        // =====================================================
        private static void AddAssemblyOccurrencesAndBodiesAsLinks(
            AssemblyDocument asmDoc,
            RobotModel robot)
        {
            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs         = asmDef.Occurrences;

            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            double scaleToMeters = 0.01;
            DebugLog("SYS",
                "AddAssemblyOccurrencesAndBodiesAsLinks: leafOccs=" + leafOccs.Count);

            int occIndex = 0;

            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    if (occ.Suppressed)
                    {
                        DebugLog("MESH",
                            "occ '" + occ.Name + "': suprimido, se omite.");
                        continue;
                    }
                    if (!occ.Visible)
                    {
                        DebugLog("MESH",
                            "occ '" + occ.Name + "': no visible, se omite.");
                        continue;
                    }

                    List<SurfaceBody> bodies = new List<SurfaceBody>();
                    CollectSurfaceBodiesFromOccurrence(occ, bodies);

                    DebugLog("MESH",
                        "AddAssemblyOccurrencesAndBodiesAsLinks: occ '" +
                        occ.Name + "', bodies=" + bodies.Count);

                    if (bodies.Count == 0)
                    {
                        DebugLog("MESH",
                            "occ '" + occ.Name +
                            "': sin SurfaceBodies/WorkSurfaces para exportar.");
                        continue;
                    }

                    Matrix m = occ.Transformation;

                    double tx_m = m.Cell[1, 4] * scaleToMeters;
                    double ty_m = m.Cell[2, 4] * scaleToMeters;
                    double tz_m = m.Cell[3, 4] * scaleToMeters;

                    double roll, pitch, yaw;
                    MatrixToRPY(m, out roll, out pitch, out yaw);

                    DebugLog("TFM",
                        "occ='" + occ.Name + "' T_world(m)=(" +
                        tx_m.ToString(CultureInfo.InvariantCulture) + ", " +
                        ty_m.ToString(CultureInfo.InvariantCulture) + ", " +
                        tz_m.ToString(CultureInfo.InvariantCulture) + ") " +
                        "rpy(rad)=(" +
                        roll.ToString(CultureInfo.InvariantCulture) + ", " +
                        pitch.ToString(CultureInfo.InvariantCulture) + ", " +
                        yaw.ToString(CultureInfo.InvariantCulture) + ")");

                    string rawName  = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    string baseLinkName = "link_" +
                                          occIndex.ToString(CultureInfo.InvariantCulture) +
                                          "_" + safeName;

                    for (int i = 0; i < bodies.Count; i++)
                    {
                        string suffix = (i == 0)
                            ? ""
                            : "_b" + i.ToString(CultureInfo.InvariantCulture);

                        string linkName = baseLinkName + suffix;

                        UrdfLink link = new UrdfLink();
                        link.Name      = linkName;
                        link.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                        link.OriginRPY = new double[] { roll, pitch, yaw };
                        robot.Links.Add(link);

                        // Por ahora todos "fixed" (0 DOF). Aquí luego
                        // se mapearán a:
                        // - revolute/prismatic → DOF del eslabón
                        // - continuous → ruedas/ejes
                        // - floating/planar → base del robot
                        UrdfJoint joint = new UrdfJoint();
                        joint.Type = "fixed";

                        if (i == 0)
                        {
                            joint.Name       = "root_" + linkName;
                            joint.ParentLink = "base_link";
                            joint.ChildLink  = linkName;
                            joint.OriginXYZ  = new double[] { tx_m, ty_m, tz_m };
                            joint.OriginRPY  = new double[] { roll, pitch, yaw };

                            DebugLog("LINK",
                                "Añadido link principal '" + linkName +
                                "' colgando de base_link.");
                        }
                        else
                        {
                            joint.Name       = "fixed_extra_" + linkName;
                            joint.ParentLink = baseLinkName;
                            joint.ChildLink  = linkName;
                            joint.OriginXYZ  = new double[] { 0.0, 0.0, 0.0 };
                            joint.OriginRPY  = new double[] { 0.0, 0.0, 0.0 };

                            DebugLog("LINK",
                                "Añadido link extra '" + linkName +
                                "' colgando de '" + baseLinkName + "'.");
                        }

                        robot.Joints.Add(joint);
                    }
                }
                catch (Exception ex)
                {
                    DebugLog("ERR",
                        "Error al crear links/joints para occurrence '" +
                        occ.Name + "': " + ex.Message);
                }
                finally
                {
                    occIndex++;
                }
            }
        }

        // =====================================================
        //  Helpers: recoger SurfaceBodies
        // =====================================================
        private static void CollectSurfaceBodiesFromPartDefinition(
            PartComponentDefinition partDef,
            List<SurfaceBody> bodies)
        {
            if (partDef == null || bodies == null)
                return;

            try
            {
                SurfaceBodies surfaceBodies = partDef.SurfaceBodies;
                if (surfaceBodies != null)
                {
                    for (int i = 1; i <= surfaceBodies.Count; i++)
                    {
                        SurfaceBody b = surfaceBodies[i];
                        if (b != null)
                            bodies.Add(b);
                    }
                }
            }
            catch { }

            try
            {
                WorkSurfaces workSurfaces = partDef.WorkSurfaces;
                if (workSurfaces != null)
                {
                    for (int wi = 1; wi <= workSurfaces.Count; wi++)
                    {
                        WorkSurface ws = workSurfaces[wi];
                        if (ws == null) continue;

                        SurfaceBodies wsBodies = ws.SurfaceBodies;
                        if (wsBodies == null) continue;

                        for (int bi = 1; bi <= wsBodies.Count; bi++)
                        {
                            SurfaceBody b2 = wsBodies[bi];
                            if (b2 != null)
                                bodies.Add(b2);
                        }
                    }
                }
            }
            catch { }
        }

        private static void CollectSurfaceBodiesFromOccurrence(
            ComponentOccurrence occ,
            List<SurfaceBody> bodies)
        {
            if (occ == null || bodies == null)
                return;

            // 1) Intentar cuerpos proxy en contexto de ensamblaje
            try
            {
                SurfaceBodies occBodies = occ.SurfaceBodies;
                if (occBodies != null && occBodies.Count > 0)
                {
                    for (int i = 1; i <= occBodies.Count; i++)
                    {
                        SurfaceBody b = occBodies[i];
                        if (b != null)
                            bodies.Add(b);
                    }
                    return;
                }
            }
            catch { }

            // 2) Fallback: cuerpos del PartDefinition
            try
            {
                PartComponentDefinition partDef = occ.Definition as PartComponentDefinition;
                if (partDef != null)
                    CollectSurfaceBodiesFromPartDefinition(partDef, bodies);
            }
            catch { }
        }

        // =====================================================
        //  MakeSafeName: limpiar nombres para URDF/DAE
        // =====================================================
        private static string MakeSafeName(string rawName)
        {
            if (string.IsNullOrEmpty(rawName))
                return "unnamed";

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < rawName.Length; i++)
            {
                char c = rawName[i];
                if ((c >= 'a' && c <= 'z') ||
                    (c >= 'A' && c <= 'Z') ||
                    (c >= '0' && c <= '9') ||
                    c == '_' || c == '-')
                {
                    sb.Append(c);
                }
                else
                {
                    sb.Append('_');
                }
            }
            return sb.ToString();
        }

        // =====================================================
        //  TESSELLATE (usa CalculateFacets)
        // =====================================================
        private static bool TessellateBodiesToMeshArrays(
            IList<SurfaceBody> bodies,
            out double[] vertices,
            out int[] indices)
        {
            vertices = null;
            indices  = null;

            if (bodies == null || bodies.Count == 0)
            {
                DebugLog("MESH", "TessellateBodiesToMeshArrays: bodies == null o Count == 0");
                return false;
            }

            List<double> vList = new List<double>();
            List<int>    iList = new List<int>();
            int vertexOffset   = 0;

            int bodyIndex = 0;
            foreach (SurfaceBody body in bodies)
            {
                if (body == null)
                {
                    DebugLog("MESH", "TessellateBodiesToMeshArrays: body[" +
                        bodyIndex.ToString(CultureInfo.InvariantCulture) + "] es null");
                    bodyIndex++;
                    continue;
                }

                DebugLog("MESH", "TessellateBodiesToMeshArrays: tessellando body[" +
                    bodyIndex.ToString(CultureInfo.InvariantCulture) + "]");
                if (!TessellateSingleBody(body, vList, iList, ref vertexOffset))
                {
                    DebugLog("MESH",
                        "TessellateBodiesToMeshArrays: body[" +
                        bodyIndex.ToString(CultureInfo.InvariantCulture) +
                        "] sin triángulos (CalculateFacets), se omite.");
                }
                bodyIndex++;
            }

            if (vList.Count == 0 || iList.Count == 0)
            {
                DebugLog("MESH",
                    "TessellateBodiesToMeshArrays: vList.Count == 0 o iList.Count == 0, no hay malla.");
                return false;
            }

            vertices = vList.ToArray();
            indices  = iList.ToArray();

            DebugLog("MESH",
                "TessellateBodiesToMeshArrays: vertsTotales=" +
                (vertices.Length / 3).ToString(CultureInfo.InvariantCulture) +
                ", trisTotales=" +
                (indices.Length / 3).ToString(CultureInfo.InvariantCulture));

            return true;
        }

        private static bool TessellateSingleBody(
            SurfaceBody body,
            List<double> vList,
            List<int> iList,
            ref int vertexOffset)
        {
            try
            {
                // Tolerancia en cm (API Inventor → cm)
                double tol = (_meshQualityMode == "high") ? 0.05 : 0.1;

                int vertexCount = 0;
                int facetCount  = 0;

                double[] vertexCoords  = new double[] { };
                double[] normalVectors = new double[] { };
                int[]    vertexIndices = new int[] { };

                body.CalculateFacets(
                    tol,
                    out vertexCount,
                    out facetCount,
                    out vertexCoords,
                    out normalVectors,
                    out vertexIndices);

                DebugLog("MESH",
                    "TessellateSingleBody: tol=" +
                    tol.ToString(CultureInfo.InvariantCulture) +
                    ", vertexCount=" + vertexCount.ToString(CultureInfo.InvariantCulture) +
                    ", facetCount="  + facetCount.ToString(CultureInfo.InvariantCulture));

                if (vertexCount <= 0 || facetCount <= 0 ||
                    vertexCoords == null || vertexCoords.Length == 0 ||
                    vertexIndices == null || vertexIndices.Length == 0)
                {
                    DebugLog("MESH", "TessellateSingleBody: CalculateFacets devolvió 0 vértices o 0 facetas.");
                    return false;
                }

                // cm → m
                for (int i = 0; i < vertexCoords.Length; i++)
                {
                    double vCm     = vertexCoords[i];
                    double vMeters = vCm * 0.01;
                    vList.Add(vMeters);
                }

                // Índices 1-based → 0-based + offset
                for (int i = 0; i < vertexIndices.Length; i++)
                {
                    int idx = vertexIndices[i] - 1;
                    if (idx < 0) idx = 0;
                    iList.Add(vertexOffset + idx);
                }

                vertexOffset = vList.Count / 3;

                DebugLog("MESH",
                    "TessellateSingleBody: vertsAcumulados=" +
                    vertexOffset.ToString(CultureInfo.InvariantCulture) +
                    ", indicesTotales=" +
                    iList.Count.ToString(CultureInfo.InvariantCulture));

                return true;
            }
            catch (Exception ex)
            {
                DebugLog("ERR", "Error en TessellateSingleBody: " + ex.Message);
                return false;
            }
        }

        // =====================================================
        //  TransformVerticesToLocalFrame
        //  (WORLD → local del componente)
        // =====================================================
        private static void TransformVerticesToLocalFrame(
            double[] verticesWorld,
            Matrix occMatrix,
            out double[] verticesLocal)
        {
            verticesLocal = null;
            if (verticesWorld == null || verticesWorld.Length == 0 || occMatrix == null)
            {
                DebugLog("TFM", "TransformVerticesToLocalFrame: sin vértices o matriz nula, no se transforma.");
                verticesLocal = verticesWorld;
                return;
            }

            double scaleToMeters = 0.01;

            // Traslación de la occurrence en metros
            double tx = occMatrix.Cell[1, 4] * scaleToMeters;
            double ty = occMatrix.Cell[2, 4] * scaleToMeters;
            double tz = occMatrix.Cell[3, 4] * scaleToMeters;

            // Rotación R
            double r11 = occMatrix.Cell[1, 1];
            double r12 = occMatrix.Cell[1, 2];
            double r13 = occMatrix.Cell[1, 3];

            double r21 = occMatrix.Cell[2, 1];
            double r22 = occMatrix.Cell[2, 2];
            double r23 = occMatrix.Cell[2, 3];

            double r31 = occMatrix.Cell[3, 1];
            double r32 = occMatrix.Cell[3, 2];
            double r33 = occMatrix.Cell[3, 3];

            // v_local = R^T * (v_world - t)
            verticesLocal = new double[verticesWorld.Length];

            for (int i = 0; i < verticesWorld.Length; i += 3)
            {
                double vx = verticesWorld[i]     - tx;
                double vy = verticesWorld[i + 1] - ty;
                double vz = verticesWorld[i + 2] - tz;

                double lx = r11 * vx + r21 * vy + r31 * vz;
                double ly = r12 * vx + r22 * vy + r32 * vz;
                double lz = r13 * vx + r23 * vy + r33 * vz;

                verticesLocal[i]     = lx;
                verticesLocal[i + 1] = ly;
                verticesLocal[i + 2] = lz;
            }

            DebugLog("TFM",
                "TransformVerticesToLocalFrame: numVerts=" +
                (verticesWorld.Length / 3).ToString(CultureInfo.InvariantCulture));
        }










        // =====================================================
        //  LOG DE ASSET / APPEARANCE (recorre TODOS los AssetValue)
        // =====================================================
        private static void LogAssetInfo(string ownerKind, string ownerName, Asset app)
        {
            if (app == null)
            {
                DebugLog("MESH",
                    "LogAssetInfo: " + ownerKind + "='" + ownerName + "' sin Asset (null).");
                return;
            }

            string appDisplayName = "(sin nombre)";
            try { appDisplayName = app.DisplayName; }
            catch { appDisplayName = "(error DisplayName)"; }

            int count = 0;
            try { count = app.Count; } catch { count = -1; }

            DebugLog("MESH",
                "LogAssetInfo: " + ownerKind +
                "='" + ownerName +
                "', Asset.DisplayName='" + appDisplayName +
                "', AssetValues: Count=" +
                count.ToString(CultureInfo.InvariantCulture));

            try
            {
                foreach (AssetValue av in app)
                {
                    if (av == null)
                    {
                        DebugLog("MESH", "    [AssetValue null]");
                        continue;
                    }

                    string avName = "";
                    string avDisplay = "";
                    bool avReadOnly = false;
                    string avType = "";

                    try { avName = av.Name; } catch { }
                    try { avDisplay = av.DisplayName; } catch { }
                    try { avReadOnly = av.IsReadOnly; } catch { }
                    try { avType = av.ValueType.ToString(); } catch { }

                    DebugLog("MESH",
                        "    AssetValue: Name='" + avName +
                        "', DisplayName='" + avDisplay +
                        "', ValueType=" + avType +
                        ", IsReadOnly=" + avReadOnly.ToString());

                    // Si es de tipo COLOR, logear también RGBA
                    try
                    {
                        if (av.ValueType == AssetValueTypeEnum.kAssetValueTypeColor)
                        {
                            ColorAssetValue cav = av as ColorAssetValue;
                            if (cav != null)
                            {
                                Inventor.Color invCol = cav.Value as Inventor.Color;
                                if (invCol != null)
                                {
                                    DebugLog("MESH",
                                        "      Color RGBA=(" +
                                        invCol.Red.ToString(CultureInfo.InvariantCulture) + "," +
                                        invCol.Green.ToString(CultureInfo.InvariantCulture) + "," +
                                        invCol.Blue.ToString(CultureInfo.InvariantCulture) + "," +
                                        invCol.Opacity.ToString(CultureInfo.InvariantCulture) + ")");
                                }
                            }
                        }
                    }
                    catch
                    {
                        DebugLog("MESH",
                            "      [Error leyendo ColorAssetValue.Value]");
                    }
                }
            }
            catch
            {
                DebugLog("MESH", "LogAssetInfo: error al iterar AssetValues.");
            }
        }

        // =====================================================
        //  Helper: buscar color por nombre de AssetValue
        //          (ej: wallpaint_color)
        // =====================================================
        private static bool TryGetColorFromNamedAssetValue(
            Asset app,
            string targetName,
            out double r,
            out double g,
            out double b)
        {
            r = 0.8;
            g = 0.8;
            b = 0.8;

            if (app == null || string.IsNullOrEmpty(targetName))
                return false;

            try
            {
                foreach (AssetValue av in app)
                {
                    if (av == null)
                        continue;

                    string avName = "";
                    try { avName = av.Name; } catch { avName = ""; }

                    if (string.IsNullOrEmpty(avName))
                        continue;

                    if (!string.Equals(avName, targetName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (av.ValueType != AssetValueTypeEnum.kAssetValueTypeColor)
                        continue;

                    ColorAssetValue cav = av as ColorAssetValue;
                    if (cav == null)
                        continue;

                    Inventor.Color invCol = cav.Value as Inventor.Color;
                    if (invCol == null)
                        continue;

                    r = invCol.Red   / 255.0;
                    g = invCol.Green / 255.0;
                    b = invCol.Blue  / 255.0;

                    DebugLog("MESH",
                        "TryGetColorFromNamedAssetValue('" + targetName +
                        "'): RGB=(" +
                        r.ToString("F3", CultureInfo.InvariantCulture) + "," +
                        g.ToString("F3", CultureInfo.InvariantCulture) + "," +
                        b.ToString("F3", CultureInfo.InvariantCulture) + ")");
                    return true;
                }
            }
            catch
            {
                DebugLog("MESH",
                    "TryGetColorFromNamedAssetValue: error buscando '" + targetName + "'.");
            }

            return false;
        }

        // =====================================================
        //  Helper central: PRIORIDAD de color dentro de un Asset
        //
        //  Orden:
        //    1) wallpaint_color
        //    2) generic_diffuse
        //    3) primer AssetValue COLOR (ej: common_tint_color)
        // =====================================================
        private static bool TryGetColorFromAssetWithPriority(
            Asset app,
            string ownerKind,
            string ownerName,
            out double r,
            out double g,
            out double b)
        {
            // Por defecto gris
            r = 0.8;
            g = 0.8;
            b = 0.8;

            if (app == null)
            {
                DebugLog("MESH",
                    "TryGetColorFromAssetWithPriority: " + ownerKind +
                    "='" + ownerName + "' sin Asset, usando gris 0.8.");
                return false;
            }

            // Log completo (para debug y para ver wallpaint/common_tint/etc.)
            LogAssetInfo(ownerKind, ownerName, app);

            // 1) PRIORIDAD: wallpaint_color (caso LEGO rojo, 178,0,0,1)
            if (TryGetColorFromNamedAssetValue(app, "wallpaint_color", out r, out g, out b))
            {
                DebugLog("MESH",
                    "TryGetColorFromAssetWithPriority: usando wallpaint_color para " +
                    ownerKind + "='" + ownerName + "'.");
                return true;
            }

            // 2) generic_diffuse, si existe
            try
            {
                AssetValue avDif = app["generic_diffuse"];
                if (avDif != null && avDif.ValueType == AssetValueTypeEnum.kAssetValueTypeColor)
                {
                    ColorAssetValue difCav = avDif as ColorAssetValue;
                    if (difCav != null)
                    {
                        Inventor.Color invCol1 = difCav.Value as Inventor.Color;
                        if (invCol1 != null)
                        {
                            r = invCol1.Red   / 255.0;
                            g = invCol1.Green / 255.0;
                            b = invCol1.Blue  / 255.0;

                            DebugLog("MESH",
                                "TryGetColorFromAssetWithPriority: " + ownerKind +
                                "='" + ownerName + "', generic_diffuse=(" +
                                r.ToString("F3", CultureInfo.InvariantCulture) + "," +
                                g.ToString("F3", CultureInfo.InvariantCulture) + "," +
                                b.ToString("F3", CultureInfo.InvariantCulture) + ")");
                            return true;
                        }
                    }
                }
            }
            catch
            {
                DebugLog("MESH",
                    "TryGetColorFromAssetWithPriority: " + ownerKind +
                    "='" + ownerName + "' sin generic_diffuse válido.");
            }

            // 3) Fallback: primer AssetValue COLOR que exista
            try
            {
                foreach (AssetValue av in app)
                {
                    if (av == null)
                        continue;

                    if (av.ValueType == AssetValueTypeEnum.kAssetValueTypeColor)
                    {
                        ColorAssetValue cav = av as ColorAssetValue;
                        if (cav != null)
                        {
                            Inventor.Color invCol = cav.Value as Inventor.Color;
                            if (invCol != null)
                            {
                                r = invCol.Red   / 255.0;
                                g = invCol.Green / 255.0;
                                b = invCol.Blue  / 255.0;

                                DebugLog("MESH",
                                    "TryGetColorFromAssetWithPriority: " + ownerKind +
                                    "='" + ownerName +
                                    "', usando primer AssetValue COLOR Name='" +
                                    av.Name + "', DisplayName='" + av.DisplayName + "', RGB=(" +
                                    r.ToString("F3", CultureInfo.InvariantCulture) + "," +
                                    g.ToString("F3", CultureInfo.InvariantCulture) + "," +
                                    b.ToString("F3", CultureInfo.InvariantCulture) + ")");
                                return true;
                            }
                        }
                    }
                }
            }
            catch
            {
                DebugLog("MESH",
                    "TryGetColorFromAssetWithPriority: error buscando AssetValue COLOR en " +
                    ownerKind + "='" + ownerName + "'.");
            }

            DebugLog("MESH",
                "TryGetColorFromAssetWithPriority: sin color, usando gris 0.8 para " +
                ownerKind + "='" + ownerName + "'.");
            return false;
        }

        // =====================================================
        //  COLOR (Body y Face) + Fallback a Occurrence.Appearance
        //
        //  Prioridad final:
        //    1) Face.Appearance (wallpaint → generic_diffuse → primer COLOR)
        //    2) Body.Appearance (wallpaint → generic_diffuse → primer COLOR)
        //    3) Occurrence.Appearance (wallpaint → generic_diffuse → primer COLOR)
        //    4) Gris (0.8, 0.8, 0.8)
        // =====================================================

        // Versión extendida con contexto de Occurrence
        private static bool TryGetBodyColor(
            SurfaceBody body,
            string ownerNameForLog,
            Asset occAppearance,
            out double r,
            out double g,
            out double b)
        {
            r = 0.8;
            g = 0.8;
            b = 0.8;

            if (body == null)
            {
                DebugLog("MESH", "TryGetBodyColor: body == null, usando gris 0.8.");
                return false;
            }

            string bodyName = "(sin nombre)";
            try
            {
                if (!string.IsNullOrEmpty(body.Name))
                    bodyName = body.Name;
            }
            catch { }

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = bodyName;

            // 1) Body.Appearance
            try
            {
                Asset appBody = null;
                try
                {
                    appBody = body.Appearance;
                }
                catch
                {
                    appBody = null;
                }

                if (appBody != null &&
                    TryGetColorFromAssetWithPriority(appBody, "Body", ownerNameForLog, out r, out g, out b))
                {
                    return true;
                }
            }
            catch
            {
                DebugLog("MESH",
                    "TryGetBodyColor: excepción leyendo Body.Appearance, se probará Occurrence.");
            }

            // 2) Fallback: Occurrence.Appearance, si nos la pasan
            if (occAppearance != null)
            {
                if (TryGetColorFromAssetWithPriority(occAppearance, "Occurrence", ownerNameForLog, out r, out g, out b))
                {
                    DebugLog("MESH",
                        "TryGetBodyColor: usando Occurrence.Appearance para body '" +
                        ownerNameForLog + "'.");
                    return true;
                }
            }

            DebugLog("MESH",
                "TryGetBodyColor: sin color detectado, usando gris 0.8 para body '" +
                ownerNameForLog + "'.");
            return false;
        }

        // Wrapper antiguo (sin Occurrence) para compatibilidad
        private static bool TryGetBodyColor(
            SurfaceBody body,
            out double r,
            out double g,
            out double b)
        {
            return TryGetBodyColor(body,
                (body != null && !string.IsNullOrEmpty(body.Name)) ? body.Name : "(body)",
                null,
                out r, out g, out b);
        }

        // Color a nivel de CARA con fallback a Body y Occurrence
        private static bool TryGetFaceColor(
            Inventor.Face face,
            SurfaceBody parentBody,
            string ownerNameForLog,
            Asset occAppearance,
            out double r,
            out double g,
            out double b)
        {
            // Por defecto gris claro
            r = 0.8;
            g = 0.8;
            b = 0.8;

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = "(Face)";

            // 1) Si hay cara, intentar Face.Appearance
            if (face != null)
            {
                try
                {
                    Asset app = null;
                    try
                    {
                        app = face.Appearance;
                    }
                    catch
                    {
                        app = null;
                    }

                    if (app != null)
                    {
                        string faceId = ownerNameForLog;
                        try
                        {
                            if (face.SurfaceBody != null && !string.IsNullOrEmpty(face.SurfaceBody.Name))
                                faceId = face.SurfaceBody.Name;
                        }
                        catch { }

                        if (TryGetColorFromAssetWithPriority(app, "Face", faceId, out r, out g, out b))
                        {
                            DebugLog("MESH",
                                "TryGetFaceColor: usando Face.Appearance para '" + faceId + "'.");
                            return true;
                        }
                    }
                }
                catch
                {
                    DebugLog("MESH",
                        "TryGetFaceColor: excepción leyendo Face.Appearance, se probará Body/Occurrence.");
                }
            }
            else
            {
                DebugLog("MESH", "TryGetFaceColor: face == null, se pasa directamente a Body/Occurrence.");
            }

            // 2) Fallback: Body (que internamente ya prueba Body.Appearance y luego Occurrence.Appearance)
            if (parentBody != null)
            {
                if (TryGetBodyColor(parentBody, ownerNameForLog, occAppearance, out r, out g, out b))
                {
                    DebugLog("MESH",
                        "TryGetFaceColor: usando color de Body/Occurrence para '" +
                        ownerNameForLog + "'.");
                    return true;
                }
            }
            else
            {
                // 3) Si no hay body pero sí Occurrence, usar Occurrence directamente
                if (occAppearance != null)
                {
                    if (TryGetColorFromAssetWithPriority(occAppearance, "Occurrence", ownerNameForLog, out r, out g, out b))
                    {
                        DebugLog("MESH",
                            "TryGetFaceColor: usando Occurrence.Appearance sin body para '" +
                            ownerNameForLog + "'.");
                        return true;
                    }
                }
            }

            DebugLog("MESH",
                "TryGetFaceColor: sin color específico, usando gris 0.8 para '" +
                ownerNameForLog + "'.");
            return false;
        }

        // Wrapper antiguo (sin Occurrence) para compatibilidad
        private static bool TryGetFaceColor(
            Inventor.Face face,
            SurfaceBody parentBody,
            out double r,
            out double g,
            out double b)
        {
            return TryGetFaceColor(
                face,
                parentBody,
                (parentBody != null && !string.IsNullOrEmpty(parentBody.Name)) ? parentBody.Name : "(body)",
                null,
                out r, out g, out b);
        }

        private static int ClampToByte(double v)
        {
            if (v < 0.0)   return 0;
            if (v > 255.0) return 255;
            return (int)Math.Round(v);
        }

        private static void WriteSolidColorPng(
            string path,
            double r,
            double g,
            double b,
            int size)
        {
            DebugLog("MESH",
                "WriteSolidColorPng: path='" + path +
                "', color=(" +
                r.ToString("F3", CultureInfo.InvariantCulture) + "," +
                g.ToString("F3", CultureInfo.InvariantCulture) + "," +
                b.ToString("F3", CultureInfo.InvariantCulture) + "), size=" +
                size.ToString(CultureInfo.InvariantCulture));

            using (Bitmap bmp = new Bitmap(size, size))
            {
                System.Drawing.Color col = System.Drawing.Color.FromArgb(
                    255,
                    ClampToByte(r * 255.0),
                    ClampToByte(g * 255.0),
                    ClampToByte(b * 255.0));

                for (int y = 0; y < size; y++)
                {
                    for (int x = 0; x < size; x++)
                    {
                        bmp.SetPixel(x, y, col);
                    }
                }

                bmp.Save(path, ImageFormat.Png);
            }
        }

        private static void WriteAtlasSingleColorPng(
            string path,
            double r,
            double g,
            double b,
            int cellsX,
            int cellsY,
            int cellSize)
        {
            int width  = cellsX * cellSize;
            int height = cellsY * cellSize;

            DebugLog("MESH",
                "WriteAtlasSingleColorPng: path='" + path +
                "', color=(" +
                r.ToString("F3", CultureInfo.InvariantCulture) + "," +
                g.ToString("F3", CultureInfo.InvariantCulture) + "," +
                b.ToString("F3", CultureInfo.InvariantCulture) + "), cellsX=" +
                cellsX.ToString(CultureInfo.InvariantCulture) + ", cellsY=" +
                cellsY.ToString(CultureInfo.InvariantCulture) + ", cellSize=" +
                cellSize.ToString(CultureInfo.InvariantCulture));

            using (Bitmap bmp = new Bitmap(width, height))
            {
                System.Drawing.Color col = System.Drawing.Color.FromArgb(
                    255,
                    ClampToByte(r * 255.0),
                    ClampToByte(g * 255.0),
                    ClampToByte(b * 255.0));

                for (int y = 0; y < height; y++)
                {
                    for (int x = 0; x < width; x++)
                    {
                        bmp.SetPixel(x, y, col);
                    }
                }

                bmp.Save(path, ImageFormat.Png);
            }
        }

        // Core del atlas con soporte para Occurrence.Appearance
        private static void WriteBodyFaceColorAtlasCore(
            SurfaceBody body,
            string ownerNameForLog,
            Asset occAppearance,
            string path,
            int cellSize)
        {
            if (body == null)
            {
                DebugLog("MESH",
                    "WriteBodyFaceColorAtlasCore: body == null, escribiendo PNG gris sólido.");
                WriteSolidColorPng(path, 0.8, 0.8, 0.8, cellSize);
                return;
            }

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = body.Name ?? "(body)";

            double bodyR, bodyG, bodyB;
            if (!TryGetBodyColor(body, ownerNameForLog, occAppearance, out bodyR, out bodyG, out bodyB))
            {
                bodyR = bodyG = bodyB = 0.8;
            }

            Faces faces = null;
            try
            {
                faces = body.Faces;
            }
            catch
            {
                faces = null;
            }

            int faceCount = (faces != null) ? faces.Count : 0;

            DebugLog("MESH",
                "WriteBodyFaceColorAtlasCore: path='" + path +
                "', faceCount=" + faceCount.ToString(CultureInfo.InvariantCulture) +
                ", owner='" + ownerNameForLog + "'");

            if (faceCount <= 0)
            {
                DebugLog("MESH",
                    "WriteBodyFaceColorAtlasCore: faceCount <= 0, usando atlas monocromático.");
                WriteAtlasSingleColorPng(path, bodyR, bodyG, bodyB, 1, 1, cellSize);
                return;
            }

            int cellsX = (int)Math.Ceiling(Math.Sqrt((double)faceCount));
            if (cellsX < 1) cellsX = 1;
            int cellsY = (int)Math.Ceiling((double)faceCount / (double)cellsX);
            if (cellsY < 1) cellsY = 1;

            int width  = cellsX * cellSize;
            int height = cellsY * cellSize;

            DebugLog("MESH",
                "WriteBodyFaceColorAtlasCore: cellsX=" +
                cellsX.ToString(CultureInfo.InvariantCulture) +
                ", cellsY=" +
                cellsY.ToString(CultureInfo.InvariantCulture) +
                ", cellSize=" +
                cellSize.ToString(CultureInfo.InvariantCulture) +
                ", width=" +
                width.ToString(CultureInfo.InvariantCulture) +
                ", height=" +
                height.ToString(CultureInfo.InvariantCulture));

            using (Bitmap bmp = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    System.Drawing.Color bgCol = System.Drawing.Color.FromArgb(
                        255,
                        ClampToByte(bodyR * 255.0),
                        ClampToByte(bodyG * 255.0),
                        ClampToByte(bodyB * 255.0));
                    g.Clear(bgCol);
                }

                for (int fi = 0; fi < faceCount; fi++)
                {
                    Inventor.Face f = null;
                    try
                    {
                        f = faces[fi + 1]; // Faces es 1-based
                    }
                    catch
                    {
                        f = null;
                    }

                    double fr, fg, fb;
                    if (!TryGetFaceColor(f, body, ownerNameForLog, occAppearance, out fr, out fg, out fb))
                    {
                        fr = bodyR;
                        fg = bodyG;
                        fb = bodyB;
                    }

                    System.Drawing.Color faceCol = System.Drawing.Color.FromArgb(
                        255,
                        ClampToByte(fr * 255.0),
                        ClampToByte(fg * 255.0),
                        ClampToByte(fb * 255.0));

                    int cellX = fi % cellsX;
                    int cellY = fi / cellsX;

                    int startX = cellX * cellSize;
                    int startY = cellY * cellSize;

                    DebugLog("MESH",
                        "WriteBodyFaceColorAtlasCore: faceIndex=" +
                        fi.ToString(CultureInfo.InvariantCulture) +
                        ", cell=(" +
                        cellX.ToString(CultureInfo.InvariantCulture) + "," +
                        cellY.ToString(CultureInfo.InvariantCulture) + "), color=(" +
                        fr.ToString("F3", CultureInfo.InvariantCulture) + "," +
                        fg.ToString("F3", CultureInfo.InvariantCulture) + "," +
                        fb.ToString("F3", CultureInfo.InvariantCulture) + ")");

                    for (int y = startY; y < startY + cellSize && y < height; y++)
                    {
                        for (int x = startX; x < startX + cellSize && x < width; x++)
                        {
                            bmp.SetPixel(x, y, faceCol);
                        }
                    }
                }

                bmp.Save(path, ImageFormat.Png);
            }

            DebugLog(
                "MESH",
                "WriteBodyFaceColorAtlasCore: atlas escrito OK en '" + path + "'");
        }

        // Wrapper antiguo (sin Occurrence) → Part .ipt
        private static void WriteBodyFaceColorAtlas(
            SurfaceBody body,
            string path,
            int cellSize)
        {
            WriteBodyFaceColorAtlasCore(
                body,
                (body != null && !string.IsNullOrEmpty(body.Name)) ? body.Name : "(body)",
                null,
                path,
                cellSize);
        }

        // Wrapper nuevo con Occurrence.Appearance → Assembly .iam
        private static void WriteBodyFaceColorAtlas(
            SurfaceBody body,
            string ownerNameForLog,
            Asset occAppearance,
            string path,
            int cellSize)
        {
            WriteBodyFaceColorAtlasCore(
                body,
                ownerNameForLog,
                occAppearance,
                path,
                cellSize);
        }
















        // =====================================================
        //  EXPORT GEOMETRY + TEXTURAS (PNG/ATLAS) POR .DAE
        // =====================================================
        private static void ExportGeometryAndTextures(
            Inventor.Application invApp,
            Document doc,
            RobotModel robot,
            string exportDir)
        {
            string meshesDir = IOPath.Combine(exportDir, "meshes");
            EnsureDirectory(meshesDir);

            DebugLog("SYS",
                "ExportGeometryAndTextures: meshesDir='" + meshesDir + "', docType=" +
                doc.DocumentType.ToString());

            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                ExportPartGeometryToDae((PartDocument)doc, robot, meshesDir);
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                ExportAssemblyGeometryToDae((AssemblyDocument)doc, robot, meshesDir);
            }
        }

        private static void ExportPartGeometryToDae(
            PartDocument partDoc,
            RobotModel robot,
            string meshesDir)
        {
            string baseName = IOPath.GetFileNameWithoutExtension(partDoc.DisplayName);
            PartComponentDefinition partDef = partDoc.ComponentDefinition;

            List<SurfaceBody> bodies = new List<SurfaceBody>();
            CollectSurfaceBodiesFromPartDefinition(partDef, bodies);

            DebugLog("MESH",
                "ExportPartGeometryToDae: Part '" + baseName +
                "': SurfaceBodies=" + bodies.Count);

            if (bodies.Count == 0)
            {
                DebugLog("MESH",
                    "ExportPartGeometryToDae: Part '" + baseName +
                    "': sin SurfaceBodies para exportar.");
            }

            for (int i = 0; i < bodies.Count; i++)
            {
                SurfaceBody body = bodies[i];
                if (body == null)
                {
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: body[" +
                        i.ToString(CultureInfo.InvariantCulture) + "] es null, se omite.");
                    continue;
                }

                string bodyName = "(null)";
                try
                {
                    if (!string.IsNullOrEmpty(body.Name))
                        bodyName = body.Name;
                }
                catch { }

                string linkName = "root_body_" +
                                  i.ToString(CultureInfo.InvariantCulture) + "_" +
                                  MakeSafeName(bodyName);

                DebugLog("MESH",
                    "ExportPartGeometryToDae: body[" +
                    i.ToString(CultureInfo.InvariantCulture) +
                    "], bodyName='" + bodyName + "', linkName='" + linkName + "'");

                UrdfLink link = FindLinkByName(robot, linkName);
                if (link == null)
                {
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: no se encontró link '" +
                        linkName + "' para body[" +
                        i.ToString(CultureInfo.InvariantCulture) + "]");
                    continue;
                }

                double[] vertices;
                int[] indices;

                List<SurfaceBody> oneBodyList = new List<SurfaceBody>();
                oneBodyList.Add(body);

                if (!TessellateBodiesToMeshArrays(oneBodyList, out vertices, out indices))
                {
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: Part '" + baseName + "', body[" +
                        i.ToString(CultureInfo.InvariantCulture) +
                        "]: tessellate no generó triángulos.");
                    continue;
                }

                string daeName  = linkName + ".dae";
                string daePath  = IOPath.Combine(meshesDir, daeName);

                DebugLog("MESH",
                    "ExportPartGeometryToDae: escribiendo DAE='" + daePath +
                    "', verts=" + (vertices.Length / 3).ToString(CultureInfo.InvariantCulture) +
                    ", tris=" + (indices.Length / 3).ToString(CultureInfo.InvariantCulture));

                WriteColladaFile(daePath, linkName, vertices, indices);

                link.MeshFile = "meshes/" + daeName;
                DebugLog("MESH",
                    "ExportPartGeometryToDae: link '" + link.Name +
                    "' MeshFile='" + link.MeshFile + "'");

                // ====== COLOR / TEXTURA PARA .IPT ======
                double r, g, b;
                if (!TryGetBodyColor(body, linkName, null, out r, out g, out b))
                {
                    r = g = b = 0.8;
                }

                string pngPath = IOPath.Combine(meshesDir, linkName + ".png");

                if (_meshQualityMode == "low")
                {
                    // VLQ → PNG sólido
                    WriteSolidColorPng(pngPath, r, g, b, 32);
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: LOW PNG='" + pngPath + "'");
                }
                else if (_meshQualityMode == "high")
                {
                    // DisplayMesh → atlas por cara
                    WriteBodyFaceColorAtlas(body, linkName, null, pngPath, 32);
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: HIGH ATLAS-PNG='" + pngPath + "'");
                }

                // ====== INERCIA ======
                try
                {
                    MassProperties mp = partDef.MassProperties;
                    FillLinkInertialFromMassProperties(link, mp);
                }
                catch
                {
                    DebugLog("SYS",
                        "ExportPartGeometryToDae: MassProperties falló para link '" +
                        link.Name + "', usando inercial dummy.");
                }
            }
        }

        // -------------------------------------------------
        //  ASSEMBLY: UN DAE + PNG/ATLAS POR BODY/LINK
        // -------------------------------------------------
        private static void ExportAssemblyGeometryToDae(
            AssemblyDocument asmDoc,
            RobotModel robot,
            string meshesDir)
        {
            if (asmDoc == null || robot == null)
            {
                DebugLog("SYS",
                    "ExportAssemblyGeometryToDae: asmDoc o robot nulos, se omite.");
                return;
            }

            // 🔧 Ajustar tipos de JOINT según constraints del .iam
            //    - InsertConstraint    → continuous (ruedas/ejes)
            //    - AngleConstraint     → revolute  (juntas rotacionales con ángulo)
            //    - TransitionalConstraint → prismatic (desplazamientos lineales)
            UpdateJointTypesFromConstraints(asmDoc, robot);

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs = asmDef.Occurrences;
            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            int occIndex = 0;

            DebugLog("SYS",
                "ExportAssemblyGeometryToDae: leafOccs=" +
                leafOccs.Count.ToString(CultureInfo.InvariantCulture) +
                ", meshesDir='" + meshesDir + "'");

            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    if (occ.Suppressed)
                    {
                        DebugLog("MESH",
                            "ExportAssemblyGeometryToDae: occ '" + occ.Name + "': suprimido, se omite.");
                        occIndex++;
                        continue;
                    }
                    if (!occ.Visible)
                    {
                        DebugLog("MESH",
                            "ExportAssemblyGeometryToDae: occ '" + occ.Name + "': no visible, se omite.");
                        occIndex++;
                        continue;
                    }

                    string rawName  = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    string baseLinkName = "link_" +
                                          occIndex.ToString(CultureInfo.InvariantCulture) +
                                          "_" + safeName;

                    // Apariencia a nivel de OCCURRENCE (para overrides de color en .iam)
                    Asset occAppearance = null;
                    try
                    {
                        occAppearance = occ.Appearance;
                    }
                    catch
                    {
                        occAppearance = null;
                    }

                    List<SurfaceBody> bodies = new List<SurfaceBody>();
                    CollectSurfaceBodiesFromOccurrence(occ, bodies);

                    DebugLog("MESH",
                        "ExportAssemblyGeometryToDae: occ '" + rawName +
                        "', bodies=" + bodies.Count.ToString(CultureInfo.InvariantCulture) +
                        ", baseLinkName='" + baseLinkName + "'");

                    if (bodies.Count == 0)
                    {
                        DebugLog("MESH",
                            "ExportAssemblyGeometryToDae: occ '" + rawName +
                            "': sin SurfaceBodies para exportar.");
                        occIndex++;
                        continue;
                    }

                    Matrix m = occ.Transformation;

                    for (int i = 0; i < bodies.Count; i++)
                    {
                        SurfaceBody body = bodies[i];
                        if (body == null)
                        {
                            DebugLog("MESH",
                                "ExportAssemblyGeometryToDae: occ '" + rawName +
                                "', body[" + i.ToString(CultureInfo.InvariantCulture) +
                                "] es null, se omite.");
                            continue;
                        }

                        string suffix = (i == 0)
                            ? ""
                            : "_b" + i.ToString(CultureInfo.InvariantCulture);

                        string linkName = baseLinkName + suffix;

                        DebugLog("MESH",
                            "ExportAssemblyGeometryToDae: occ '" + rawName +
                            "', body[" + i.ToString(CultureInfo.InvariantCulture) +
                            "], linkName='" + linkName + "'");

                        UrdfLink link = FindLinkByName(robot, linkName);
                        if (link == null)
                        {
                            DebugLog("MESH",
                                "ExportAssemblyGeometryToDae: occ '" + rawName +
                                "', body[" + i.ToString(CultureInfo.InvariantCulture) +
                                "]: no hay link '" + linkName + "', se omite.");
                            continue;
                        }

                        double[] verticesWorld;
                        int[] indices;

                        List<SurfaceBody> oneBodyList = new List<SurfaceBody>();
                        oneBodyList.Add(body);

                        if (!TessellateBodiesToMeshArrays(oneBodyList, out verticesWorld, out indices))
                        {
                            DebugLog("MESH",
                                "ExportAssemblyGeometryToDae: occ '" + rawName +
                                "', body[" + i.ToString(CultureInfo.InvariantCulture) +
                                "]: tessellate no generó triángulos.");
                            continue;
                        }

                        double[] verticesLocal;
                        TransformVerticesToLocalFrame(verticesWorld, m, out verticesLocal);

                        string daeName = linkName + ".dae";
                        string daePath = IOPath.Combine(meshesDir, daeName);

                        DebugLog("MESH",
                            "ExportAssemblyGeometryToDae: occ '" + rawName +
                            "', body[" + i.ToString(CultureInfo.InvariantCulture) +
                            "]: DAE='" + daePath +
                            "', verts=" + (verticesLocal.Length / 3).ToString(CultureInfo.InvariantCulture) +
                            ", tris=" + (indices.Length / 3).ToString(CultureInfo.InvariantCulture));

                        WriteColladaFile(daePath, linkName, verticesLocal, indices);

                        link.MeshFile = "meshes/" + daeName;
                        DebugLog("MESH",
                            "ExportAssemblyGeometryToDae: link '" + link.Name +
                            "' MeshFile='" + link.MeshFile + "'");

                        // ====== COLOR / TEXTURA PARA .IAM ======
                        // Usamos prioridad:
                        //   Face.Appearance → Body.Appearance → Occurrence.Appearance → gris
                        // (la parte de Face vs Body vs Occurrence está dentro de TryGetFaceColor/TryGetBodyColor,
                        //  y dentro de cada Asset: wallpaint_color → generic_diffuse → primer COLOR)
                        double r, g, b;
                        if (!TryGetBodyColor(body, rawName, occAppearance, out r, out g, out b))
                        {
                            r = g = b = 0.8;
                        }

                        string pngPath = IOPath.Combine(meshesDir, linkName + ".png");

                        if (_meshQualityMode == "low")
                        {
                            // VLQ → PNG sólido del color efectivo de esta occurrence/body
                            WriteSolidColorPng(pngPath, r, g, b, 32);
                            DebugLog("MESH",
                                "ExportAssemblyGeometryToDae: LOW PNG='" + pngPath + "'");
                        }
                        else if (_meshQualityMode == "high")
                        {
                            // DisplayMesh → atlas por cara usando Face/Body/Occurrence
                            WriteBodyFaceColorAtlas(body, rawName, occAppearance, pngPath, 32);
                            DebugLog("MESH",
                                "ExportAssemblyGeometryToDae: HIGH ATLAS-PNG='" + pngPath + "'");
                        }

                        // ====== INERCIA ======
                        try
                        {
                            PartComponentDefinition partDef = occ.Definition as PartComponentDefinition;
                            if (partDef != null)
                            {
                                MassProperties mp = partDef.MassProperties;
                                FillLinkInertialFromMassProperties(link, mp);
                            }
                            else
                            {
                                DebugLog("SYS",
                                    "ExportAssemblyGeometryToDae: occ '" + rawName +
                                    "' sin PartComponentDefinition, inercial dummy para link '" +
                                    link.Name + "'.");
                            }
                        }
                        catch
                        {
                            DebugLog("SYS",
                                "ExportAssemblyGeometryToDae: MassProperties falló para occ '" +
                                rawName + "', link '" + link.Name + "'.");
                        }
                    }

                    occIndex++;
                }
                catch (Exception ex)
                {
                    DebugLog("ERR",
                        "ExportAssemblyGeometryToDae: Error al exportar geometría para occ '" +
                        occ.Name + "': " + ex.Message);
                }
            }
        }












        // -------------------------------------------------
        //  Buscar link por nombre
        // -------------------------------------------------
        private static UrdfLink FindLinkByName(RobotModel robot, string name)
        {
            if (robot == null || robot.Links == null)
                return null;

            foreach (UrdfLink link in robot.Links)
            {
                if (link != null && link.Name == name)
                    return link;
            }
            return null;
        }

        // -------------------------------------------------
        //  Buscar JOINT por ChildLink (para mapear constraints)
        // -------------------------------------------------
        private static UrdfJoint FindJointByChildLink(RobotModel robot, string childLinkName)
        {
            if (robot == null || robot.Joints == null)
                return null;
            if (string.IsNullOrEmpty(childLinkName))
                return null;

            foreach (UrdfJoint j in robot.Joints)
            {
                if (j != null && j.ChildLink == childLinkName)
                    return j;
            }
            return null;
        }

        // -------------------------------------------------
        //  Mapear AssemblyConstraints → tipos de JOINT URDF
        // -------------------------------------------------
        private static void UpdateJointTypesFromConstraints(
            AssemblyDocument asmDoc,
            RobotModel robot)
        {
            if (asmDoc == null || robot == null)
            {
                DebugLog("LINK",
                    "UpdateJointTypesFromConstraints: asmDoc o robot nulos, se omite.");
                return;
            }

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            if (asmDef == null)
            {
                DebugLog("LINK",
                    "UpdateJointTypesFromConstraints: asmDef nulo, se omite.");
                return;
            }

            ComponentOccurrences occs = asmDef.Occurrences;
            if (occs == null)
            {
                DebugLog("LINK",
                    "UpdateJointTypesFromConstraints: asmDef.Occurrences nulo, se omite.");
                return;
            }

            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            // Mapeo robusto: ComponentOccurrence → nombre baseLink
            Dictionary<ComponentOccurrence, string> occToBaseLink =
                new Dictionary<ComponentOccurrence, string>();

            int occIndex = 0;
            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    if (occ == null)
                    {
                        occIndex++;
                        continue;
                    }

                    if (occ.Suppressed || !occ.Visible)
                    {
                        occIndex++;
                        continue;
                    }

                    string rawName  = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    string baseLinkName = "link_" +
                        occIndex.ToString(CultureInfo.InvariantCulture) +
                        "_" + safeName;

                    occToBaseLink[occ] = baseLinkName;

                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: occIndex=" +
                        occIndex.ToString(CultureInfo.InvariantCulture) +
                        ", occName='" + rawName +
                        "', baseLinkName='" + baseLinkName + "'");
                }
                catch (Exception ex)
                {
                    DebugLog("ERR",
                        "UpdateJointTypesFromConstraints: error mapeando occurrence[" +
                        occIndex.ToString(CultureInfo.InvariantCulture) + "]: " + ex.Message);
                }
                finally
                {
                    occIndex++;
                }
            }

            AssemblyConstraints constraints = null;
            try
            {
                constraints = asmDef.Constraints;
            }
            catch
            {
                constraints = null;
            }

            if (constraints == null || constraints.Count == 0)
            {
                DebugLog("LINK",
                    "UpdateJointTypesFromConstraints: sin AssemblyConstraints en el ensamblaje.");
                return;
            }

            DebugLog("LINK",
                "UpdateJointTypesFromConstraints: constraints.Count=" +
                constraints.Count.ToString(CultureInfo.InvariantCulture));

            foreach (AssemblyConstraint ac in constraints)
            {
                if (ac == null)
                    continue;

                ComponentOccurrence childOcc = null;
                ComponentOccurrence parentOcc = null;

                try
                {
                    // Convención simple: OccurrenceOne → CHILD, OccurrenceTwo → PARENT
                    childOcc = ac.OccurrenceOne;
                    parentOcc = ac.OccurrenceTwo;
                }
                catch
                {
                    childOcc = null;
                    parentOcc = null;
                }

                if (childOcc == null)
                    continue;

                string baseLinkName;
                if (!occToBaseLink.TryGetValue(childOcc, out baseLinkName))
                {
                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: constraint '" + ac.Name +
                        "' usa occurrenceOne que no está en occToBaseLink (suprimido o no hoja).");
                    continue;
                }

                UrdfJoint joint = FindJointByChildLink(robot, baseLinkName);
                if (joint == null)
                {
                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: no se encontró JOINT para childLink='" +
                        baseLinkName + "' (constraint '" + ac.Name + "').");
                    continue;
                }

                // Sólo reasignar si estaba como FIXED (para no pisar algo que se cambió antes)
                if (!string.Equals(joint.Type, "fixed", StringComparison.OrdinalIgnoreCase))
                {
                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: joint '" + joint.Name +
                        "' ya es type='" + joint.Type + "', no se modifica desde constraint '" +
                        ac.Name + "'.");
                    continue;
                }

                string newType = null;
                double[] newAxis = null;

                // -------- Mapeo básico por tipo de constraint --------
                if (ac is InsertConstraint)
                {
                    // InsertConstraint: típico eje rotacional (rueda/pin)
                    newType = "continuous"; // rotación infinita
                    newAxis = new double[] { 0.0, 0.0, 1.0 };
                }
                else if (ac is AngleConstraint)
                {
                    // AngleConstraint: rotación limitada (aunque aquí no leemos los límites)
                    newType = "revolute";
                    newAxis = new double[] { 0.0, 0.0, 1.0 };
                }
                else if (ac is TransitionalConstraint)
                {
                    // TransitionalConstraint: desplazamiento lineal entre perfiles → prismatic
                    newType = "prismatic";
                    newAxis = new double[] { 0.0, 0.0, 1.0 };
                }
                else
                {
                    // Otros tipos (Mate, Flush, Tangent, etc.) se mantienen como fixed,
                    // porque normalmente solo fijan o restringen, no definen un DOF explícito.
                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: constraint '" + ac.Name +
                        "' no mapeado a DOF (se deja joint FIXED). ObjectType=" +
                        ac.Type.ToString());
                }

                if (!string.IsNullOrEmpty(newType))
                {
                    joint.Type = newType;
                    if (newAxis != null && newAxis.Length == 3)
                        joint.AxisXYZ = newAxis;

                    string parentLinkName = joint.ParentLink ?? "base_link";
                    string parentNameInfo = (parentOcc != null) ? parentOcc.Name : "(parent desconocido)";

                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: constraint '" + ac.Name +
                        "' → joint '" + joint.Name +
                        "', parentLink='" + parentLinkName +
                        "', childLink='" + joint.ChildLink +
                        "', mappedType='" + newType +
                        "', parentOcc='" + parentNameInfo + "'");
                }
            }
        }

        // -------------------------------------------------
        //  COLLADA (DAE) con referencia a textura PNG
        // -------------------------------------------------
        private static void WriteColladaFile(
            string fullPath,
            string geometryName,
            double[] vertices,
            int[] indices)
        {
            DebugLog("MESH",
                "WriteColladaFile: fullPath='" + fullPath +
                "', geometryName='" + geometryName +
                "', numVerts=" + (vertices != null ? (vertices.Length / 3).ToString(CultureInfo.InvariantCulture) : "0") +
                ", numTris=" + (indices != null ? (indices.Length / 3).ToString(CultureInfo.InvariantCulture) : "0"));

            string text = BuildColladaText(geometryName, vertices, indices);
            IOFile.WriteAllText(fullPath, text);
        }

        private static string BuildColladaText(
            string geometryName,
            double[] vertices,
            int[] indices)
        {
            if (vertices == null) vertices = new double[0];
            if (indices  == null) indices  = new int[0];

            StringBuilder sb = new StringBuilder();

            string geomId            = geometryName + "-geom";
            string positionsId       = geometryName + "-positions";
            string positionsArrayId  = positionsId + "-array";
            string verticesId        = geometryName + "-verts";

            string imageId           = geometryName + "-image";
            string effectId          = geometryName + "-effect";
            string materialId        = geometryName + "-material";
            string surfaceSid        = geometryName + "-surface";
            string samplerSid        = geometryName + "-sampler";

            string textureFileName   = geometryName + ".png";

            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine("<COLLADA xmlns=\"http://www.collada.org/2005/11/COLLADASchema\" version=\"1.4.1\">");
            sb.AppendLine("  <asset>");
            sb.AppendLine("    <contributor>");
            sb.AppendLine("      <authoring_tool>URDFConverterAddIn</authoring_tool>");
            sb.AppendLine("    </contributor>");
            sb.AppendLine("    <unit name=\"meter\" meter=\"1\"/>");
            sb.AppendLine("    <up_axis>Z_UP</up_axis>");
            sb.AppendLine("  </asset>");

            // IMAGEN (para que figure el path de textura en el DAE)
            sb.AppendLine("  <library_images>");
            sb.AppendLine("    <image id=\"" + imageId + "\" name=\"" + imageId + "\">");
            sb.AppendLine("      <init_from>" + textureFileName + "</init_from>");
            sb.AppendLine("    </image>");
            sb.AppendLine("  </library_images>");

            // EFFECT (usa la imagen como textura difusa)
            sb.AppendLine("  <library_effects>");
            sb.AppendLine("    <effect id=\"" + effectId + "\">");
            sb.AppendLine("      <profile_COMMON>");
            sb.AppendLine("        <newparam sid=\"" + surfaceSid + "\">");
            sb.AppendLine("          <surface type=\"2D\">");
            sb.AppendLine("            <init_from>" + imageId + "</init_from>");
            sb.AppendLine("          </surface>");
            sb.AppendLine("        </newparam>");
            sb.AppendLine("        <newparam sid=\"" + samplerSid + "\">");
            sb.AppendLine("          <sampler2D>");
            sb.AppendLine("            <source>" + surfaceSid + "</source>");
            sb.AppendLine("          </sampler2D>");
            sb.AppendLine("        </newparam>");
            sb.AppendLine("        <technique sid=\"common\">");
            sb.AppendLine("          <lambert>");
            sb.AppendLine("            <diffuse>");
            sb.AppendLine("              <texture texture=\"" + samplerSid + "\" texcoord=\"TEX0\"/>");
            sb.AppendLine("            </diffuse>");
            sb.AppendLine("          </lambert>");
            sb.AppendLine("        </technique>");
            sb.AppendLine("      </profile_COMMON>");
            sb.AppendLine("    </effect>");
            sb.AppendLine("  </library_effects>");

            // MATERIAL
            sb.AppendLine("  <library_materials>");
            sb.AppendLine("    <material id=\"" + materialId + "\" name=\"" + materialId + "\">");
            sb.AppendLine("      <instance_effect url=\"#" + effectId + "\"/>");
            sb.AppendLine("    </material>");
            sb.AppendLine("  </library_materials>");

            // GEOMETRÍA
            sb.AppendLine("  <library_geometries>");
            sb.AppendLine("    <geometry id=\"" + geomId + "\" name=\"" + geometryName + "\">");
            sb.AppendLine("      <mesh>");

            // Positions
            sb.AppendLine("        <source id=\"" + positionsId + "\">");
            sb.Append("          <float_array id=\"")
              .Append(positionsArrayId)
              .Append("\" count=\"")
              .Append(vertices.Length.ToString(CultureInfo.InvariantCulture))
              .Append("\">");

            for (int i = 0; i < vertices.Length; i++)
            {
                sb.Append(FloatToString(vertices[i]));
                if (i + 1 < vertices.Length)
                    sb.Append(" ");
            }
            sb.AppendLine("</float_array>");
            sb.AppendLine("          <technique_common>");
            sb.AppendLine("            <accessor source=\"#" + positionsArrayId + "\" count=\"" + (vertices.Length / 3).ToString(CultureInfo.InvariantCulture) + "\" stride=\"3\">");
            sb.AppendLine("              <param name=\"X\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"Y\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"Z\" type=\"float\"/>");
            sb.AppendLine("            </accessor>");
            sb.AppendLine("          </technique_common>");
            sb.AppendLine("        </source>");

            sb.AppendLine("        <vertices id=\"" + verticesId + "\">");
            sb.AppendLine("          <input semantic=\"POSITION\" source=\"#" + positionsId + "\"/>");
            sb.AppendLine("        </vertices>");


            int triCount = indices.Length / 3;
            sb.AppendLine("        <triangles material=\"" + materialId + "\" count=\"" + triCount.ToString(CultureInfo.InvariantCulture) + "\">");
            sb.AppendLine("          <input semantic=\"VERTEX\" source=\"#" + verticesId + "\" offset=\"0\"/>");
            sb.Append("          <p>");
            for (int i = 0; i < indices.Length; i++)
            {
                sb.Append(indices[i].ToString(CultureInfo.InvariantCulture));
                if (i + 1 < indices.Length)
                    sb.Append(" ");
            }
            sb.AppendLine("</p>");
            sb.AppendLine("        </triangles>");

            sb.AppendLine("      </mesh>");
            sb.AppendLine("    </geometry>");
            sb.AppendLine("  </library_geometries>");

            sb.AppendLine("  <library_visual_scenes>");
            sb.AppendLine("    <visual_scene id=\"Scene\" name=\"Scene\">");
            sb.AppendLine("      <node id=\"" + geometryName + "_node\" name=\"" + geometryName + "\">");
            sb.AppendLine("        <instance_geometry url=\"#" + geomId + "\">");
            sb.AppendLine("          <bind_material>");
            sb.AppendLine("            <technique_common>");
            sb.AppendLine("              <instance_material symbol=\"" + materialId + "\" target=\"#" + materialId + "\"/>");
            sb.AppendLine("            </technique_common>");
            sb.AppendLine("          </bind_material>");
            sb.AppendLine("        </instance_geometry>");
            sb.AppendLine("      </node>");
            sb.AppendLine("    </visual_scene>");
            sb.AppendLine("  </library_visual_scenes>");

            sb.AppendLine("  <scene>");
            sb.AppendLine("    <instance_visual_scene url=\"#Scene\"/>");
            sb.AppendLine("  </scene>");
            sb.AppendLine("</COLLADA>");

            return sb.ToString();
        }

        private static string FloatToString(double value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        // =====================================================
        //  INERTIAL DESDE MassProperties
        // =====================================================
        private static void FillLinkInertialFromMassProperties(
            UrdfLink link,
            MassProperties mp)
        {
            if (link == null || mp == null)
            {
                DebugLog("SYS", "FillLinkInertialFromMassProperties: link o mp nulos, se omite.");
                return;
            }

            double mass = mp.Mass;

            Inventor.Point com = mp.CenterOfMass;
            double scaleToMeters = 0.01;
            double comGlobalX = com.X * scaleToMeters;
            double comGlobalY = com.Y * scaleToMeters;
            double comGlobalZ = com.Z * scaleToMeters;

            double originX = 0.0;
            double originY = 0.0;
            double originZ = 0.0;
            if (link.OriginXYZ != null && link.OriginXYZ.Length == 3)
            {
                originX = link.OriginXYZ[0];
                originY = link.OriginXYZ[1];
                originZ = link.OriginXYZ[2];
            }

            double comLocalX = comGlobalX - originX;
            double comLocalY = comGlobalY - originY;
            double comLocalZ = comGlobalZ - originZ;

            double Ixx, Iyy, Izz, Ixy, Iyz, Ixz;
            mp.XYZMomentsOfInertia(out Ixx, out Iyy, out Izz, out Ixy, out Iyz, out Ixz);

            double inertiaScale = scaleToMeters * scaleToMeters;
            Ixx *= inertiaScale;
            Iyy *= inertiaScale;
            Izz *= inertiaScale;
            Ixy *= inertiaScale;
            Iyz *= inertiaScale;
            Ixz *= inertiaScale;

            link.HasInertial = true;
            link.Mass = mass;

            link.InertialOriginXYZ = new double[]
            {
                comLocalX, comLocalY, comLocalZ
            };
            link.InertialOriginRPY = new double[]
            {
                0.0, 0.0, 0.0
            };

            link.Ixx = Ixx;
            link.Iyy = Iyy;
            link.Izz = Izz;
            link.Ixy = Ixy;
            link.Iyz = Iyz;
            link.Ixz = Ixz;

            DebugLog(
                "SYS",
                "Inercial link '" + link.Name + "': mass=" +
                mass.ToString("G5", CultureInfo.InvariantCulture) +
                " kg, COM_local=(" +
                comLocalX.ToString("G5", CultureInfo.InvariantCulture) + "," +
                comLocalY.ToString("G5", CultureInfo.InvariantCulture) + "," +
                comLocalZ.ToString("G5", CultureInfo.InvariantCulture) + ") " +
                "I=(Ixx=" + Ixx.ToString("G5", CultureInfo.InvariantCulture) +
                ", Iyy=" + Iyy.ToString("G5", CultureInfo.InvariantCulture) +
                ", Izz=" + Izz.ToString("G5", CultureInfo.InvariantCulture) + ")");
        }

        // =====================================================
        //  WriteUrdfFile
        // =====================================================
        private static void WriteUrdfFile(RobotModel robot, string urdfPath)
        {
            if (robot == null)
            {
                DebugLog("SYS", "WriteUrdfFile: robot == null, no se escribe URDF.");
                return;
            }

            DebugLog("SYS",
                "WriteUrdfFile: urdfPath='" + urdfPath +
                "', numLinks=" + robot.Links.Count.ToString(CultureInfo.InvariantCulture) +
                ", numJoints=" + robot.Joints.Count.ToString(CultureInfo.InvariantCulture));

            StringBuilder sb = new StringBuilder();

            string robotName = robot.Name;
            if (string.IsNullOrEmpty(robotName))
                robotName = "InventorRobot";

            sb.AppendLine("<?xml version=\"1.0\"?>");
            sb.AppendLine("<robot name=\"" + XmlEscape(robotName) + "\">");

            // LINKS
            foreach (UrdfLink link in robot.Links)
            {
                if (link == null)
                    continue;

                DebugLog("SYS",
                    "WriteUrdfFile: LINK name='" + link.Name +
                    "', mesh='" + (link.MeshFile ?? "(null)") +
                    "', hasInertial=" + link.HasInertial.ToString());

                sb.AppendLine("  <link name=\"" + XmlEscape(link.Name) + "\">");

                // INERTIAL
                if (link.HasInertial)
                {
                    string xyzIn = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} {1} {2}",
                        link.InertialOriginXYZ[0],
                        link.InertialOriginXYZ[1],
                        link.InertialOriginXYZ[2]);

                    string rpyIn = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} {1} {2}",
                        link.InertialOriginRPY[0],
                        link.InertialOriginRPY[1],
                        link.InertialOriginRPY[2]);

                    sb.AppendLine("    <inertial>");
                    sb.AppendLine("      <origin xyz=\"" + xyzIn + "\" rpy=\"" + rpyIn + "\"/>");
                    sb.AppendLine("      <mass value=\"" +
                        link.Mass.ToString(CultureInfo.InvariantCulture) + "\"/>");
                    sb.AppendLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "      <inertia ixx=\"{0}\" ixy=\"{1}\" ixz=\"{2}\" iyy=\"{3}\" iyz=\"{4}\" izz=\"{5}\"/>",
                        link.Ixx, link.Ixy, link.Ixz, link.Iyy, link.Iyz, link.Izz));
                    sb.AppendLine("    </inertial>");
                }
                else
                {
                    sb.AppendLine("    <inertial>");
                    sb.AppendLine("      <origin xyz=\"0 0 0\" rpy=\"0 0 0\"/>");
                    sb.AppendLine("      <mass value=\"1e-06\"/>");
                    sb.AppendLine("      <inertia ixx=\"1e-06\" ixy=\"0\" ixz=\"0\" iyy=\"1e-06\" iyz=\"0\" izz=\"1e-06\"/>");
                    sb.AppendLine("    </inertial>");
                }

                // VISUAL
                if (!string.IsNullOrEmpty(link.MeshFile))
                {
                    sb.AppendLine("    <visual>");
                    sb.AppendLine("      <origin xyz=\"0 0 0\" rpy=\"0 0 0\"/>");
                    sb.AppendLine("      <geometry>");
                    sb.AppendLine("        <mesh filename=\"" + XmlEscape(link.MeshFile) + "\"/>");
                    sb.AppendLine("      </geometry>");
                    sb.AppendLine("    </visual>");
                }

                sb.AppendLine("  </link>");
            }

            // JOINTS
            foreach (UrdfJoint joint in robot.Joints)
            {
                if (joint == null)
                    continue;

                DebugLog("SYS",
                    "WriteUrdfFile: JOINT name='" + joint.Name +
                    "', type='" + joint.Type +
                    "', parent='" + joint.ParentLink +
                    "', child='" + joint.ChildLink + "'");

                sb.AppendLine("  <joint name=\"" +
                    XmlEscape(joint.Name) + "\" type=\"" +
                    XmlEscape(joint.Type) + "\">");

                sb.AppendLine("    <parent link=\"" +
                    XmlEscape(joint.ParentLink) + "\"/>");
                sb.AppendLine("    <child link=\"" +
                    XmlEscape(joint.ChildLink) + "\"/>");

                string xyz = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0} {1} {2}",
                    joint.OriginXYZ[0],
                    joint.OriginXYZ[1],
                    joint.OriginXYZ[2]);

                string rpy = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0} {1} {2}",
                    joint.OriginRPY[0],
                    joint.OriginRPY[1],
                    joint.OriginRPY[2]);

                sb.AppendLine("    <origin xyz=\"" + xyz + "\" rpy=\"" + rpy + "\"/>");

                bool isFixed      = string.Equals(joint.Type, "fixed", StringComparison.OrdinalIgnoreCase);
                bool isRevolute   = string.Equals(joint.Type, "revolute", StringComparison.OrdinalIgnoreCase);
                bool isPrismatic  = string.Equals(joint.Type, "prismatic", StringComparison.OrdinalIgnoreCase);
                bool isContinuous = string.Equals(joint.Type, "continuous", StringComparison.OrdinalIgnoreCase);

                // Eje del joint (para revolute/prismatic/continuous)
                if (!isFixed && joint.AxisXYZ != null && joint.AxisXYZ.Length == 3)
                {
                    string axisStr = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} {1} {2}",
                        joint.AxisXYZ[0],
                        joint.AxisXYZ[1],
                        joint.AxisXYZ[2]);
                    sb.AppendLine("    <axis xyz=\"" + axisStr + "\"/>");
                }

                // Límites básicos (por defecto) para revolute/prismatic
                if (isRevolute || isPrismatic)
                {
                    sb.AppendLine("    <limit lower=\"-3.14159\" upper=\"3.14159\" effort=\"1\" velocity=\"1\"/>");
                }

                sb.AppendLine("  </joint>");
            }

            sb.AppendLine("</robot>");

            IOFile.WriteAllText(urdfPath, sb.ToString());

            DebugLog("SYS", "WriteUrdfFile: URDF guardado en '" + urdfPath + "'");
        }

        // =====================================================
        //  XmlEscape
        // =====================================================
        private static string XmlEscape(string s)
        {
            if (s == null)
                return "";

            string result = s;
            result = result.Replace("&", "&amp;");
            result = result.Replace("<", "&lt;");
            result = result.Replace(">", "&gt;");
            result = result.Replace("\"", "&quot;");
            return result;
        }

        // =====================================================
        //  MatrixToRPY
        // =====================================================
        private static void MatrixToRPY(Matrix m, out double roll, out double pitch, out double yaw)
        {
            double r11 = m.Cell[1, 1];
            double r12 = m.Cell[1, 2];
            double r13 = m.Cell[1, 3];

            double r21 = m.Cell[2, 1];
            double r22 = m.Cell[2, 2];
            double r23 = m.Cell[2, 3];

            double r31 = m.Cell[3, 1];
            double r32 = m.Cell[3, 2];
            double r33 = m.Cell[3, 3];

            double sy = Math.Sqrt(r11 * r11 + r21 * r21);
            bool singular = sy < 1e-6;

            if (!singular)
            {
                pitch = Math.Atan2(-r31, sy);
                roll  = Math.Atan2(r32, r33);
                yaw   = Math.Atan2(r21, r11);
            }
            else
            {
                pitch = Math.Atan2(-r31, sy);
                roll  = 0.0;
                yaw   = Math.Atan2(-r12, r22);
            }

            DebugLog("TFM",
                "MatrixToRPY: roll=" + roll.ToString("F4", CultureInfo.InvariantCulture) +
                ", pitch=" + pitch.ToString("F4", CultureInfo.InvariantCulture) +
                ", yaw=" + yaw.ToString("F4", CultureInfo.InvariantCulture));
        }

    } // fin clase UrdfExporter


    // =====================================================
    //  Clases simples de modelo URDF
    // =====================================================
    public class RobotModel
    {
        public string Name;

        public List<UrdfLink> Links;
        public List<UrdfJoint> Joints;

        public RobotModel()
        {
            Links  = new List<UrdfLink>();
            Joints = new List<UrdfJoint>();
        }
    }

    public class UrdfLink
    {
        public string Name;

        public string MeshFile;

        public double[] OriginXYZ;
        public double[] OriginRPY;

        public bool HasInertial = false;

        public double Mass = 1e-6;

        public double[] InertialOriginXYZ = new double[] { 0.0, 0.0, 0.0 };
        public double[] InertialOriginRPY = new double[] { 0.0, 0.0, 0.0 };

        public double Ixx = 1e-6;
        public double Iyy = 1e-6;
        public double Izz = 1e-6;
        public double Ixy = 0.0;
        public double Iyz = 0.0;
        public double Ixz = 0.0;
    }

    /*
     * Tipos clásicos de joints en URDF (1–6) y usos típicos:
     *
     * 1) Revolute:
     *    - Articulación rotacional con límites (mínimo/máximo).
     *    - Es uno de los dos tipos más usados en robots seriales (brazos manipuladores),
     *      porque aporta un grado de libertad rotacional (DOF) para orientar el efector final.
     *
     * 2) Prismatic:
     *    - Articulación lineal con límites, que se desplaza a lo largo de un eje.
     *    - Junto con Revolute, es el otro tipo más usado en robots seriales,
     *      aportando un DOF traslacional para mover o posicionar el efector final.
     *
     * 3) Continuous:
     *    - Rotación infinita sin límites (como una bisagra 360°).
     *    - Ideal para ruedas, ejes que giran libremente o juntas que deben rotar sin fin.
     *
     * 4) Fixed:
     *    - Sin grados de libertad (0 DOF).
     *    - Fundamental para simplificar la cinemática cuando sólo queremos "pegar"
     *      un modelo visual o de colisión a un eslabón principal sin movimiento relativo.
     *
     * 5) Floating:
     *    - 6 DOF (3 traslacionales + 3 rotacionales).
     *    - Se usa casi siempre para definir el joint de la base del robot respecto
     *      al marco de referencia del mundo (world frame) en la simulación, permitiendo
     *      que el robot se mueva libremente en el espacio.
     *
     * 6) Planar:
     *    - 3 DOF (2 traslaciones + 1 rotación) restringidos a un plano.
     *    - También se utiliza frecuentemente para modelar la base del robot sobre un plano
     *      (por ejemplo, un robot móvil sobre el suelo), permitiendo movimiento libre en ese plano.
     */

    public class UrdfJoint
    {
        public string Name;
        public string Type;
        public string ParentLink;
        public string ChildLink;

        public double[] OriginXYZ;
        public double[] OriginRPY;

        // Eje del joint (para revolute / prismatic / continuous)
        public double[] AxisXYZ;

        public UrdfJoint()
        {
            Type = "fixed";
            OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
            OriginRPY = new double[] { 0.0, 0.0, 0.0 };
            AxisXYZ   = new double[] { 0.0, 0.0, 1.0 }; // Z por defecto
        }
    }

} // fin namespace URDFConverterAddIn












