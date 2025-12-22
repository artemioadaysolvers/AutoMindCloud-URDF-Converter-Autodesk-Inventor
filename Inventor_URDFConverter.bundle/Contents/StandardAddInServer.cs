// ===============================
//  BLOQUE 1/4  (CÓDIGO ARREGLADO)
//  - Filtrado robusto de constraints (evita joints falsos)
//  - Selección parent/child más correcta
//  - Extracción de eje MÁS robusta (incluye Circle/Arc, y punto si existe)
//  - NO convierte a joint si no hay eje útil (Angle/Transitional/Mate/Insert)
//  - MateConstraint: solo si la fuente del eje es “cilíndrica / axis-like”
// ===============================

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using System.Reflection;   // <-- AQUI
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;               // Para PNGs
using System.Drawing.Imaging;       // Para guardar PNG
using System.IO;
using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using Inventor;
using DrawingPoint = System.Drawing.Point;
using InvPoint = Inventor.Point;

namespace URDFConverterAddIn
{
    // GUID NUEVO: 13c0f7be-eb12-48e9-963a-83e672efe557
    [Guid("13c0f7be-eb12-48e9-963a-83e672efe557")]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application _invApp;

        // Dos botones: VeryLowOptimized y DisplayMesh
        private ButtonDefinition _exportUrdfVlqButton;
        private ButtonDefinition _exportUrdfDisplayButton;

        // ClientId del AddIn: DEBE ser el mismo GUID que arriba, pero con llaves
        private const string AddInClientId = "{13c0f7be-eb12-48e9-963a-83e672efe557}";

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
                _exportUrdfVlqButton = null;
                try { _exportUrdfVlqButton = controlDefs["urdf_export_vlq_cmd"] as ButtonDefinition; }
                catch (Exception exLookup)
                {
                    Debug.WriteLine("[URDF][SYS] lookup 'urdf_export_vlq_cmd' lanzó: " + exLookup.Message);
                    _exportUrdfVlqButton = null;
                }

                if (_exportUrdfVlqButton == null)
                {
                    _exportUrdfVlqButton = controlDefs.AddButtonDefinition(
                        "Export URDF (VLQ)",
                        "urdf_export_vlq_cmd",
                        CommandTypesEnum.kNonShapeEditCmdType,
                        AddInClientId,
                        "Export URDF with VeryLowOptimized mesh",
                        "Export URDF (VLQ)");
                }

                try { _exportUrdfVlqButton.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfVlqButtonPressed); } catch { }
                _exportUrdfVlqButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfVlqButtonPressed);

                _exportUrdfDisplayButton = null;
                try { _exportUrdfDisplayButton = controlDefs["urdf_export_display_cmd"] as ButtonDefinition; }
                catch (Exception exLookup)
                {
                    Debug.WriteLine("[URDF][SYS] lookup 'urdf_export_display_cmd' lanzó: " + exLookup.Message);
                    _exportUrdfDisplayButton = null;
                }

                if (_exportUrdfDisplayButton == null)
                {
                    _exportUrdfDisplayButton = controlDefs.AddButtonDefinition(
                        "Export URDF (Display)",
                        "urdf_export_display_cmd",
                        CommandTypesEnum.kNonShapeEditCmdType,
                        AddInClientId,
                        "Export URDF with DisplayMesh-quality mesh",
                        "Export URDF (Display)");
                }

                try { _exportUrdfDisplayButton.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed); } catch { }
                _exportUrdfDisplayButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed);

                Debug.WriteLine("[URDF][SYS] ButtonDefinitions creados/obtenidos correctamente.");

                // -------------------------------------------------
                // 2) Añadir los botones a los Ribbons de Part y Assembly
                // -------------------------------------------------
                UserInterfaceManager uiMgr = _invApp.UserInterfaceManager;

                Debug.WriteLine("[URDF][SYS] Creando panels en ribbons...");

                try
                {
                    Ribbon partRibbon = uiMgr.Ribbons["Part"];
                    RibbonTab toolsTabPart = partRibbon.RibbonTabs["id_TabTools"];

                    RibbonPanel urdfPanelPart = null;
                    try { urdfPanelPart = toolsTabPart.RibbonPanels["urdf_export_panel_part"]; }
                    catch { urdfPanelPart = null; }

                    if (urdfPanelPart == null)
                    {
                        urdfPanelPart = toolsTabPart.RibbonPanels.Add(
                            "URDF Export",
                            "urdf_export_panel_part",
                            AddInClientId,
                            "",
                            false);
                    }

                    SafeAddButtonToPanel(urdfPanelPart, _exportUrdfVlqButton);
                    SafeAddButtonToPanel(urdfPanelPart, _exportUrdfDisplayButton);

                    Debug.WriteLine("[URDF][UI] Panel URDF en Part creado/actualizado correctamente.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[URDF][UI] Error creando panel en Part: " + ex.Message);
                }

                try
                {
                    Ribbon asmRibbon = uiMgr.Ribbons["Assembly"];
                    RibbonTab toolsTabAsm = asmRibbon.RibbonTabs["id_TabTools"];

                    RibbonPanel urdfPanelAsm = null;
                    try { urdfPanelAsm = toolsTabAsm.RibbonPanels["urdf_export_panel_asm"]; }
                    catch { urdfPanelAsm = null; }

                    if (urdfPanelAsm == null)
                    {
                        urdfPanelAsm = toolsTabAsm.RibbonPanels.Add(
                            "URDF Export",
                            "urdf_export_panel_asm",
                            AddInClientId,
                            "",
                            false);
                    }

                    SafeAddButtonToPanel(urdfPanelAsm, _exportUrdfVlqButton);
                    SafeAddButtonToPanel(urdfPanelAsm, _exportUrdfDisplayButton);

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

        private static void SafeAddButtonToPanel(RibbonPanel panel, ButtonDefinition btn)
        {
            if (panel == null || btn == null) return;

            try
            {
                CommandControls ctrls = panel.CommandControls;
                if (ctrls != null)
                {
                    for (int i = 1; i <= ctrls.Count; i++)
                    {
                        try
                        {
                            CommandControl cc = ctrls[i];
                            if (cc == null) continue;
                            if (cc.ControlDefinition == null) continue;

                            string internalName = "";
                            try { internalName = cc.ControlDefinition.InternalName; } catch { internalName = ""; }

                            if (!string.IsNullOrEmpty(internalName) &&
                                string.Equals(internalName, btn.InternalName, StringComparison.OrdinalIgnoreCase))
                            {
                                return; // ya está
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            try { panel.CommandControls.AddButton(btn); }
            catch { }
        }

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
                    try { _exportUrdfVlqButton.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfVlqButtonPressed); } catch { }
                    Marshal.FinalReleaseComObject(_exportUrdfVlqButton);
                    _exportUrdfVlqButton = null;
                }

                if (_exportUrdfDisplayButton != null)
                {
                    try { _exportUrdfDisplayButton.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed); } catch { }
                    Marshal.FinalReleaseComObject(_exportUrdfDisplayButton);
                    _exportUrdfDisplayButton = null;
                }
            }
            catch { }

            if (_invApp != null)
            {
                try { Marshal.ReleaseComObject(_invApp); } catch { }
                _invApp = null;
            }
        }

        public void ExecuteCommand(int CommandID) { }
        public object Automation { get { return null; } }
    }

    // ========================================================================
    //  URDF EXPORTER (BLOQUE 1/4: base + builder links/joints + JOINT MAPPING FIX)
    // ========================================================================
    public static class UrdfExporter
    {
        // "low"  → VeryLowOptimized
        // "high" → DisplayMesh
        private static string _meshQualityMode = "low";

        public static void SetMeshQualityVeryLow() { _meshQualityMode = "low"; }
        public static void SetMeshQualityDisplay() { _meshQualityMode = "high"; }
        public static string GetMeshQualityMode() { return _meshQualityMode; }

        // Debug flags
        private static bool _DEBUG_SYS = true;
        private static bool _DEBUG_TRANSFORMS = true;
        private static bool _DEBUG_MESH_TREE = true;
        private static bool _DEBUG_LINK_JOINT = true;

        private static void DebugLog(string category, string message)
        {
            if (category == "SYS" && !_DEBUG_SYS) return;
            if (category == "TFM" && !_DEBUG_TRANSFORMS) return;
            if (category == "MESH" && !_DEBUG_MESH_TREE) return;
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
            try { fullPath = doc.FullFileName; } catch { fullPath = string.Empty; }

            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show("El documento no tiene ruta de fichero. Guárdalo antes de exportar.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string baseDir = IOPath.GetDirectoryName(fullPath);
            string baseName = IOPath.GetFileNameWithoutExtension(fullPath);

            DebugLog("SYS",
                "ExportActiveDocument: doc='" + doc.DisplayName +
                "', type=" + doc.DocumentType.ToString() +
                ", path='" + fullPath +
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
            DebugLog("SYS", "ExportActiveDocument: exportDir='" + exportDir + "', urdfPath='" + urdfPath + "'");

            try
            {
                // 1) Construir el modelo URDF (links + joints base)
                RobotModel robot = BuildRobotFromDocument(doc, baseName);

                // 2) Exportar geometría + PNG/Atlas por .dae  (BLOQUE 2/4)
                ExportGeometryAndTextures(invApp, doc, robot, exportDir);

                // 3) Escribir .urdf (BLOQUE 4/4)
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

            DebugLog("SYS", "Robot construido: links=" + robot.Links.Count + ", joints=" + robot.Joints.Count);
            return robot;
        }

        // =====================================================
        //  PART: un link por SurfaceBody
        // =====================================================
        private static void AddPartBodiesAsLinks(PartDocument partDoc, RobotModel robot, string baseName)
        {
            PartComponentDefinition partDef = partDoc.ComponentDefinition;

            List<SurfaceBody> bodies = new List<SurfaceBody>();
            CollectSurfaceBodiesFromPartDefinition(partDef, bodies);

            DebugLog("MESH", "AddPartBodiesAsLinks: part='" + baseName + "', SurfaceBodies=" + bodies.Count);

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

                DebugLog("LINK", "Part sin SurfaceBodies: creado link único '" + linkName + "'");
                return;
            }

            for (int i = 0; i < bodies.Count; i++)
            {
                SurfaceBody b = bodies[i];
                string bodyName = "(null)";
                try { if (b != null && !string.IsNullOrEmpty(b.Name)) bodyName = b.Name; } catch { }

                string linkName = "root_body_" + i.ToString(CultureInfo.InvariantCulture) + "_" + MakeSafeName(bodyName);

                UrdfLink link2 = new UrdfLink();
                link2.Name = linkName;
                link2.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                link2.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                robot.Links.Add(link2);

                UrdfJoint joint2 = new UrdfJoint();
                joint2.Name = "root_" + linkName;
                joint2.Type = "fixed";
                joint2.ParentLink = "base_link";
                joint2.ChildLink = linkName;
                joint2.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                joint2.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                robot.Joints.Add(joint2);

                DebugLog("LINK", "AddPartBodiesAsLinks: creado link '" + linkName + "' para SurfaceBody[" + i.ToString(CultureInfo.InvariantCulture) + "]");
            }
        }

        // =====================================================
        //  ASSEMBLY: links por occurrence hoja (+ bodies _bN)
        // =====================================================
        private static void AddAssemblyOccurrencesAndBodiesAsLinks(AssemblyDocument asmDoc, RobotModel robot)
        {
            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs = asmDef.Occurrences;
            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            double scaleToMeters = 0.01;

            DebugLog("SYS", "AddAssemblyOccurrencesAndBodiesAsLinks: leafOccs=" + leafOccs.Count);

            int occIndex = 0;

            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    if (occ.Suppressed)
                    {
                        DebugLog("MESH", "occ '" + occ.Name + "': suprimido, se omite.");
                        continue;
                    }
                    if (!occ.Visible)
                    {
                        DebugLog("MESH", "occ '" + occ.Name + "': no visible, se omite.");
                        continue;
                    }

                    List<SurfaceBody> bodies = new List<SurfaceBody>();
                    CollectSurfaceBodiesFromOccurrence(occ, bodies);

                    DebugLog("MESH", "AddAssemblyOccurrencesAndBodiesAsLinks: occ '" + occ.Name + "', bodies=" + bodies.Count);

                    if (bodies.Count == 0)
                    {
                        DebugLog("MESH", "occ '" + occ.Name + "': sin SurfaceBodies/WorkSurfaces para exportar.");
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
                        tz_m.ToString(CultureInfo.InvariantCulture) + ") rpy(rad)=(" +
                        roll.ToString(CultureInfo.InvariantCulture) + ", " +
                        pitch.ToString(CultureInfo.InvariantCulture) + ", " +
                        yaw.ToString(CultureInfo.InvariantCulture) + ")");

                    string safeName = MakeSafeName(occ.Name);

                    string baseLinkName = "link_" + occIndex.ToString(CultureInfo.InvariantCulture) + "_" + safeName;

                    for (int i = 0; i < bodies.Count; i++)
                    {
                        string suffix = (i == 0) ? "" : "_b" + i.ToString(CultureInfo.InvariantCulture);
                        string linkName = baseLinkName + suffix;

                        UrdfLink link = new UrdfLink();
                        link.Name = linkName;
                        link.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                        link.OriginRPY = new double[] { roll, pitch, yaw };
                        robot.Links.Add(link);

                        UrdfJoint joint = new UrdfJoint();
                        joint.Type = "fixed";

                        if (i == 0)
                        {
                            joint.Name = "root_" + linkName;
                            joint.ParentLink = "base_link";
                            joint.ChildLink = linkName;
                            joint.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                            joint.OriginRPY = new double[] { roll, pitch, yaw };
                            DebugLog("LINK", "Añadido link principal '" + linkName + "' colgando de base_link.");
                        }
                        else
                        {
                            joint.Name = "fixed_extra_" + linkName;
                            joint.ParentLink = baseLinkName;
                            joint.ChildLink = linkName;
                            joint.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                            joint.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
                            DebugLog("LINK", "Añadido link extra '" + linkName + "' colgando de '" + baseLinkName + "'.");
                        }

                        robot.Joints.Add(joint);
                    }
                }
                catch (Exception ex)
                {
                    DebugLog("LINK", "Error al crear links/joints para occurrence '" + occ.Name + "': " + ex.Message);
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
        private static void CollectSurfaceBodiesFromPartDefinition(PartComponentDefinition partDef, List<SurfaceBody> bodies)
        {
            if (partDef == null || bodies == null) return;

            try
            {
                SurfaceBodies surfaceBodies = partDef.SurfaceBodies;
                if (surfaceBodies != null)
                {
                    for (int i = 1; i <= surfaceBodies.Count; i++)
                    {
                        SurfaceBody b = surfaceBodies[i];
                        if (b != null) bodies.Add(b);
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
                            if (b2 != null) bodies.Add(b2);
                        }
                    }
                }
            }
            catch { }
        }

        private static void CollectSurfaceBodiesFromOccurrence(ComponentOccurrence occ, List<SurfaceBody> bodies)
        {
            if (occ == null || bodies == null) return;

            try
            {
                SurfaceBodies occBodies = occ.SurfaceBodies;
                if (occBodies != null && occBodies.Count > 0)
                {
                    for (int i = 1; i <= occBodies.Count; i++)
                    {
                        SurfaceBody b = occBodies[i];
                        if (b != null) bodies.Add(b);
                    }
                    return;
                }
            }
            catch { }

            try
            {
                PartComponentDefinition partDef = occ.Definition as PartComponentDefinition;
                if (partDef != null)
                    CollectSurfaceBodiesFromPartDefinition(partDef, bodies);
            }
            catch { }
        }

        // =====================================================
        //  MakeSafeName
        // =====================================================
        private static string MakeSafeName(string rawName)
        {
            if (string.IsNullOrEmpty(rawName)) return "unnamed";

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

        // =====================================================================
        //  ====== FIX: MAPEO ROBUSTO DE CONSTRAINTS → JOINTS (EVITA JOINTS FALSOS)
        // =====================================================================

        private static UrdfJoint FindJointByChildLink(RobotModel robot, string childLinkName)
        {
            if (robot == null || robot.Joints == null) return null;
            if (string.IsNullOrEmpty(childLinkName)) return null;

            foreach (UrdfJoint j in robot.Joints)
                if (j != null && j.ChildLink == childLinkName)
                    return j;

            return null;
        }

        private static ComponentOccurrence ResolveToMappedLeafOccurrence(
            ComponentOccurrence occ,
            Dictionary<ComponentOccurrence, string> occToBaseLink)
        {
            if (occ == null) return null;
            if (occToBaseLink != null && occToBaseLink.ContainsKey(occ))
                return occ;

            try
            {
                ComponentOccurrencesEnumerator subs = null;
                try { subs = occ.SubOccurrences; } catch { subs = null; }

                if (subs != null)
                {
                    foreach (ComponentOccurrence so in subs)
                    {
                        if (so == null) continue;

                        if (occToBaseLink != null && occToBaseLink.ContainsKey(so))
                            return so;

                        ComponentOccurrence deep = ResolveToMappedLeafOccurrence(so, occToBaseLink);
                        if (deep != null && occToBaseLink != null && occToBaseLink.ContainsKey(deep))
                            return deep;
                    }
                }
            }
            catch { }

            return occ;
        }

        private static bool IsAxisSourceCylindricalOrAxisLike(string axisSrc)
        {
            if (string.IsNullOrEmpty(axisSrc)) return false;
            string s = axisSrc.ToLowerInvariant();

            // Axis-like real: cilindro/cone/circle/arc/line/workaxis/edge(line)
            if (s.Contains("cylinder")) return true;
            if (s.Contains("cone")) return true;
            if (s.Contains("circle")) return true;
            if (s.Contains("arc")) return true;
            if (s.Contains("workaxis")) return true;
            if (s.Contains("line.direction")) return true;
            if (s.Contains("linesegment.direction")) return true;
            if (s.Contains("edge.geometry")) return true;
            return false;
        }

        private static bool ConstraintImpliesMovableJoint(AssemblyConstraint ac, bool gotAxis, string axisSrc, out string urdfType)
        {
            urdfType = null;
            if (ac == null) return false;

            // Nunca crear joint móvil si NO hay eje (evita 90% de falsos positivos)
            // (Insert/Angle/Transitional/Mate pueden existir “para alinear” y quedar rígidos)
            if (!gotAxis) return false;

            if (ac is InsertConstraint)
            {
                // Insert casi siempre define eje real (perno). Requiere eje.
                urdfType = "continuous";
                return true;
            }
            if (ac is TransitionalConstraint)
            {
                // Prismatic requiere eje.
                urdfType = "prismatic";
                return true;
            }
            if (ac is AngleConstraint)
            {
                // Angle puede ser “solo orientación” entre planos; si no tenemos eje útil, no convertir.
                urdfType = "revolute";
                return true;
            }
            if (ac is MateConstraint)
            {
                // Mate es el gran generador de joints falsos:
                // solo aceptamos si el eje viene de geometría cilíndrica/axis-like.
                if (!IsAxisSourceCylindricalOrAxisLike(axisSrc))
                    return false;

                urdfType = "continuous";
                return true;
            }

            return false;
        }

        // -------------------------------------------------
        //  Mapear AssemblyConstraints → tipos de JOINT URDF (ROBUSTO)
        //  - NO pisa joints ya definidos (solo cambia fixed→movible)
        //  - Selección parent/child más correcta
        //  - NO convierte si no hay eje (y en Mate requiere axis-like)
        // -------------------------------------------------
        private static void UpdateJointTypesFromConstraints(AssemblyDocument asmDoc, RobotModel robot)
        {
            if (asmDoc == null || robot == null) return;

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            if (asmDef == null) return;

            ComponentOccurrences occs = asmDef.Occurrences;
            if (occs == null) return;

            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            Dictionary<ComponentOccurrence, string> occToBaseLink = new Dictionary<ComponentOccurrence, string>();
            Dictionary<ComponentOccurrence, Matrix> occToMatrix = new Dictionary<ComponentOccurrence, Matrix>();

            int occIndex = 0;
            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    if (occ == null) { occIndex++; continue; }
                    if (occ.Suppressed || !occ.Visible) { occIndex++; continue; }

                    string safeName = MakeSafeName(occ.Name);
                    string baseLinkName = "link_" + occIndex.ToString(CultureInfo.InvariantCulture) + "_" + safeName;

                    occToBaseLink[occ] = baseLinkName;
                    try { occToMatrix[occ] = occ.Transformation; } catch { }
                }
                catch { }
                finally { occIndex++; }
            }

            AssemblyConstraints constraints = null;
            try { constraints = asmDef.Constraints; } catch { constraints = null; }
            if (constraints == null || constraints.Count == 0) return;

            foreach (AssemblyConstraint ac in constraints)
            {
                if (ac == null) continue;

                // Skip constraints suprimidas si existe la propiedad
                try
                {
                    bool sup = false;
                    try { sup = ac.Suppressed; } catch { sup = false; }
                    if (sup) continue;
                }
                catch { }

                ComponentOccurrence o1 = null, o2 = null;
                try { o1 = ac.OccurrenceOne; o2 = ac.OccurrenceTwo; } catch { o1 = null; o2 = null; }

                ComponentOccurrence ro1 = ResolveToMappedLeafOccurrence(o1, occToBaseLink);
                ComponentOccurrence ro2 = ResolveToMappedLeafOccurrence(o2, occToBaseLink);

                string link1 = null, link2 = null;
                if (ro1 != null) occToBaseLink.TryGetValue(ro1, out link1);
                if (ro2 != null) occToBaseLink.TryGetValue(ro2, out link2);

                if (string.IsNullOrEmpty(link1) && string.IsNullOrEmpty(link2))
                    continue;

                // Elegir child/parent robusto:
                //  - grounded -> parent
                //  - sino: el que actualmente cuelga fixed se vuelve child preferido
                ComponentOccurrence childOcc = null;
                ComponentOccurrence parentOcc = null;
                string childLink = null;
                string parentLink = null;

                bool g1 = false, g2 = false;
                try { if (ro1 != null) g1 = ro1.Grounded; } catch { g1 = false; }
                try { if (ro2 != null) g2 = ro2.Grounded; } catch { g2 = false; }

                if (!string.IsNullOrEmpty(link1) && !string.IsNullOrEmpty(link2))
                {
                    if (g1 && !g2) { parentOcc = ro1; childOcc = ro2; parentLink = link1; childLink = link2; }
                    else if (g2 && !g1) { parentOcc = ro2; childOcc = ro1; parentLink = link2; childLink = link1; }
                    else
                    {
                        UrdfJoint j1 = FindJointByChildLink(robot, link1);
                        UrdfJoint j2 = FindJointByChildLink(robot, link2);

                        bool j1Fixed = (j1 != null && string.Equals(j1.Type, "fixed", StringComparison.OrdinalIgnoreCase));
                        bool j2Fixed = (j2 != null && string.Equals(j2.Type, "fixed", StringComparison.OrdinalIgnoreCase));

                        // Preferir como child el que hoy está fijo (lo estamos “liberando”)
                        if (j2Fixed && !j1Fixed)
                        {
                            parentOcc = ro1; childOcc = ro2; parentLink = link1; childLink = link2;
                        }
                        else
                        {
                            parentOcc = ro2; childOcc = ro1; parentLink = link2; childLink = link1;
                        }
                    }
                }
                else
                {
                    // Constraint con algo no mapeado (sub-ensamble): usamos el mapeado como child
                    if (!string.IsNullOrEmpty(link1))
                    {
                        childOcc = ro1; parentOcc = ro2; childLink = link1; parentLink = link2;
                    }
                    else
                    {
                        childOcc = ro2; parentOcc = ro1; childLink = link2; parentLink = link1;
                    }
                }

                if (childOcc == null || string.IsNullOrEmpty(childLink))
                    continue;

                UrdfJoint joint = FindJointByChildLink(robot, childLink);
                if (joint == null) continue;

                // Solo reasignar si estaba FIXED (no pisar joints ya definidos)
                if (!string.Equals(joint.Type, "fixed", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Extraer eje desde el constraint
                double[] axisWorldDir = null;
                double[] axisWorldPoint = null; // (por ahora sólo se captura; el pivot se usará en Bloque 2/4+ si quieres)
                string axisSrc = null;
                bool gotAxis = TryExtractAxisWorldFromConstraint(ac, out axisWorldDir, out axisWorldPoint, out axisSrc);

                // Decide si realmente corresponde a un joint móvil
                string newType;
                if (!ConstraintImpliesMovableJoint(ac, gotAxis, axisSrc, out newType))
                    continue;

                joint.Type = newType;

                // Parent real si lo tenemos
                if (!string.IsNullOrEmpty(parentLink))
                {
                    if (!string.Equals(parentLink, joint.ChildLink, StringComparison.OrdinalIgnoreCase))
                        joint.ParentLink = parentLink;
                }

                // Origin relativo parent→child (mantiene consistencia con meshes en frame del occurrence)
                try
                {
                    Matrix parentM = null, childM = null;
                    if (parentOcc != null) occToMatrix.TryGetValue(parentOcc, out parentM);
                    if (childOcc != null) occToMatrix.TryGetValue(childOcc, out childM);

                    double tx_m, ty_m, tz_m, rr, pp, yy;
                    if (TryComputeRelativeXYZRPY(parentM, childM, out tx_m, out ty_m, out tz_m, out rr, out pp, out yy))
                    {
                        joint.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                        joint.OriginRPY = new double[] { rr, pp, yy };
                    }
                }
                catch { }

                // Axis local (frame del CHILD link, consistente con tu pipeline actual)
                if (!string.Equals(joint.Type, "fixed", StringComparison.OrdinalIgnoreCase) &&
                    gotAxis && axisWorldDir != null && axisWorldDir.Length == 3)
                {
                    Matrix childM = null;
                    occToMatrix.TryGetValue(childOcc, out childM);

                    double[] axisLocal = AxisWorldToOccLocal(axisWorldDir, childM);
                    if (axisLocal != null && axisLocal.Length == 3)
                        joint.AxisXYZ = axisLocal;
                }

                DebugLog("LINK",
                    "Constraint→Joint: type=" + joint.Type +
                    ", parent='" + joint.ParentLink + "', child='" + joint.ChildLink +
                    "', axisSrc='" + (axisSrc ?? "") + "'" +
                    (joint.AxisXYZ != null ? (", axisLocal=(" +
                        joint.AxisXYZ[0].ToString("F3", CultureInfo.InvariantCulture) + "," +
                        joint.AxisXYZ[1].ToString("F3", CultureInfo.InvariantCulture) + "," +
                        joint.AxisXYZ[2].ToString("F3", CultureInfo.InvariantCulture) + ")") : ", axisLocal=(null)"));
            }
        }

        // =====================================================
        //  RELATIVE origin: parent→child (m) + RPY(rad)
        // =====================================================
        private static bool TryComputeRelativeXYZRPY(
            Matrix parentM,
            Matrix childM,
            out double tx_m,
            out double ty_m,
            out double tz_m,
            out double roll,
            out double pitch,
            out double yaw)
        {
            tx_m = ty_m = tz_m = 0.0;
            roll = pitch = yaw = 0.0;

            // Si parentM es null, asumimos base_link identidad
            if (childM == null) return false;

            if (parentM == null)
            {
                tx_m = childM.Cell[1, 4] * 0.01;
                ty_m = childM.Cell[2, 4] * 0.01;
                tz_m = childM.Cell[3, 4] * 0.01;
                MatrixToRPY(childM, out roll, out pitch, out yaw);
                return true;
            }

            // Parent R
            double pr11 = parentM.Cell[1, 1], pr12 = parentM.Cell[1, 2], pr13 = parentM.Cell[1, 3];
            double pr21 = parentM.Cell[2, 1], pr22 = parentM.Cell[2, 2], pr23 = parentM.Cell[2, 3];
            double pr31 = parentM.Cell[3, 1], pr32 = parentM.Cell[3, 2], pr33 = parentM.Cell[3, 3];

            // Child R
            double cr11 = childM.Cell[1, 1], cr12 = childM.Cell[1, 2], cr13 = childM.Cell[1, 3];
            double cr21 = childM.Cell[2, 1], cr22 = childM.Cell[2, 2], cr23 = childM.Cell[2, 3];
            double cr31 = childM.Cell[3, 1], cr32 = childM.Cell[3, 2], cr33 = childM.Cell[3, 3];

            // R_rel = Rp^T * Rc
            double r11 = pr11 * cr11 + pr21 * cr21 + pr31 * cr31;
            double r12 = pr11 * cr12 + pr21 * cr22 + pr31 * cr32;
            double r13 = pr11 * cr13 + pr21 * cr23 + pr31 * cr33;

            double r21 = pr12 * cr11 + pr22 * cr21 + pr32 * cr31;
            double r22 = pr12 * cr12 + pr22 * cr22 + pr32 * cr32;
            double r23 = pr12 * cr13 + pr22 * cr23 + pr32 * cr33;

            double r31 = pr13 * cr11 + pr23 * cr21 + pr33 * cr31;
            double r32 = pr13 * cr12 + pr23 * cr22 + pr33 * cr32;
            double r33 = pr13 * cr13 + pr23 * cr23 + pr33 * cr33;

            // t_rel(cm) = Rp^T * (tc - tp)
            double dtx_cm = childM.Cell[1, 4] - parentM.Cell[1, 4];
            double dty_cm = childM.Cell[2, 4] - parentM.Cell[2, 4];
            double dtz_cm = childM.Cell[3, 4] - parentM.Cell[3, 4];

            double tx_cm = pr11 * dtx_cm + pr21 * dty_cm + pr31 * dtz_cm;
            double ty_cm = pr12 * dtx_cm + pr22 * dty_cm + pr32 * dtz_cm;
            double tz_cm = pr13 * dtx_cm + pr23 * dty_cm + pr33 * dtz_cm;

            tx_m = tx_cm * 0.01;
            ty_m = ty_cm * 0.01;
            tz_m = tz_cm * 0.01;

            // R_rel -> RPY
            double sy = Math.Sqrt(r11 * r11 + r21 * r21);
            bool singular = sy < 1e-6;

            if (!singular)
            {
                pitch = Math.Atan2(-r31, sy);
                roll = Math.Atan2(r32, r33);
                yaw = Math.Atan2(r21, r11);
            }
            else
            {
                pitch = Math.Atan2(-r31, sy);
                roll = 0.0;
                yaw = Math.Atan2(-r12, r22);
            }

            return true;
        }

        // =====================================================
        //  EXTRAER AXIS WORLD (+ punto si se puede) DESDE CONSTRAINT
        // =====================================================
        private static bool TryExtractAxisWorldFromConstraint(
            AssemblyConstraint ac,
            out double[] axisWorldDir,
            out double[] axisWorldPoint,
            out string axisSource)
        {
            axisWorldDir = null;
            axisWorldPoint = null;
            axisSource = null;
            if (ac == null) return false;

            object g1 = null, g2 = null, e1 = null, e2 = null;

            try
            {
                if (ac is MateConstraint)
                {
                    MateConstraint mc = (MateConstraint)ac;
                    try { g1 = mc.GeometryOne; } catch { }
                    try { g2 = mc.GeometryTwo; } catch { }
                    try { e1 = mc.EntityOne; } catch { }
                    try { e2 = mc.EntityTwo; } catch { }
                }
                else if (ac is InsertConstraint)
                {
                    InsertConstraint ic = (InsertConstraint)ac;
                    try { g1 = ic.GeometryOne; } catch { }
                    try { g2 = ic.GeometryTwo; } catch { }
                    try { e1 = ic.EntityOne; } catch { }
                    try { e2 = ic.EntityTwo; } catch { }
                }
                else if (ac is AngleConstraint)
                {
                    AngleConstraint ang = (AngleConstraint)ac;
                    try { g1 = ang.GeometryOne; } catch { }
                    try { g2 = ang.GeometryTwo; } catch { }
                    try { e1 = ang.EntityOne; } catch { }
                    try { e2 = ang.EntityTwo; } catch { }
                }
                else if (ac is TransitionalConstraint)
                {
                    TransitionalConstraint tc = (TransitionalConstraint)ac;
                    try { g1 = tc.GeometryOne; } catch { }
                    try { g2 = tc.GeometryTwo; } catch { }
                    try { e1 = tc.EntityOne; } catch { }
                    try { e2 = tc.EntityTwo; } catch { }
                }
                else
                {
                    try { e1 = ac.EntityOne; } catch { }
                    try { e2 = ac.EntityTwo; } catch { }
                    try { g1 = ac.GeometryOne; } catch { }
                    try { g2 = ac.GeometryTwo; } catch { }
                }
            }
            catch { }

            if (TryGetAxisFromAnyObject(g1, out axisWorldDir, out axisWorldPoint, out axisSource)) return true;
            if (TryGetAxisFromAnyObject(g2, out axisWorldDir, out axisWorldPoint, out axisSource)) return true;
            if (TryGetAxisFromAnyObject(e1, out axisWorldDir, out axisWorldPoint, out axisSource)) return true;
            if (TryGetAxisFromAnyObject(e2, out axisWorldDir, out axisWorldPoint, out axisSource)) return true;

            return false;
        }

        // Devuelve:
        //  axisWorldDir: vector normalizado (en coords del objeto que devuelve Inventor; normalmente assembly/world)
        //  axisWorldPoint: un punto sobre el eje si está disponible (puede ser null)
        private static bool TryGetAxisFromAnyObject(
            object obj,
            out double[] axisWorldDir,
            out double[] axisWorldPoint,
            out string source)
        {
            axisWorldDir = null;
            axisWorldPoint = null;
            source = null;
            if (obj == null) return false;

            try
            {
                Cylinder cyl = obj as Cylinder;
                if (cyl != null)
                {
                    UnitVector uv = cyl.AxisVector;
                    axisWorldDir = Normalize3(new double[] { uv.X, uv.Y, uv.Z });

                    InvPoint bp = null;
                    try { bp = cyl.BasePoint; } catch { bp = null; }
                    if (bp != null) axisWorldPoint = new double[] { bp.X, bp.Y, bp.Z };

                    source = "Cylinder.AxisVector";
                    return axisWorldDir != null;
                }

                Cone cone = obj as Cone;
                if (cone != null)
                {
                    UnitVector uv = cone.AxisVector;
                    axisWorldDir = Normalize3(new double[] { uv.X, uv.Y, uv.Z });

                    InvPoint bp = null;
                    try { bp = cone.BasePoint; } catch { bp = null; }
                    if (bp != null) axisWorldPoint = new double[] { bp.X, bp.Y, bp.Z };

                    source = "Cone.AxisVector";
                    return axisWorldDir != null;
                }

                Circle circle = obj as Circle;
                if (circle != null)
                {
                    UnitVector n = circle.Normal;
                    axisWorldDir = Normalize3(new double[] { n.X, n.Y, n.Z });

                    InvPoint c = null;
                    try { c = circle.Center; } catch { c = null; }
                    if (c != null) axisWorldPoint = new double[] { c.X, c.Y, c.Z };

                    source = "Circle.Normal";
                    return axisWorldDir != null;
                }

                Arc3d arc = obj as Arc3d;
                if (arc != null)
                {
                    UnitVector n = arc.Normal;
                    axisWorldDir = Normalize3(new double[] { n.X, n.Y, n.Z });

                    InvPoint c = null;
                    try { c = arc.Center; } catch { c = null; }
                    if (c != null) axisWorldPoint = new double[] { c.X, c.Y, c.Z };

                    source = "Arc3d.Normal";
                    return axisWorldDir != null;
                }

                Line line = obj as Line;
                if (line != null)
                {
                    UnitVector uv = line.Direction;
                    axisWorldDir = Normalize3(new double[] { uv.X, uv.Y, uv.Z });

                    InvPoint rp = null;
                    try { rp = line.RootPoint; } catch { rp = null; }
                    if (rp != null) axisWorldPoint = new double[] { rp.X, rp.Y, rp.Z };

                    source = "Line.Direction";
                    return axisWorldDir != null;
                }

                LineSegment seg = obj as LineSegment;
                if (seg != null)
                {
                    UnitVector uv = seg.Direction;
                    axisWorldDir = Normalize3(new double[] { uv.X, uv.Y, uv.Z });

                    InvPoint sp = null;
                    try { sp = seg.StartPoint; } catch { sp = null; }
                    if (sp != null) axisWorldPoint = new double[] { sp.X, sp.Y, sp.Z };

                    source = "LineSegment.Direction";
                    return axisWorldDir != null;
                }

                WorkAxis wa = obj as WorkAxis;
                if (wa != null)
                {
                    Line wl = null;
                    try { wl = wa.Line; } catch { wl = null; }
                    if (wl != null)
                    {
                        UnitVector uv = wl.Direction;
                        axisWorldDir = Normalize3(new double[] { uv.X, uv.Y, uv.Z });

                        InvPoint rp = null;
                        try { rp = wl.RootPoint; } catch { rp = null; }
                        if (rp != null) axisWorldPoint = new double[] { rp.X, rp.Y, rp.Z };

                        source = "WorkAxis.Line.Direction";
                        return axisWorldDir != null;
                    }
                }

                Edge ed = obj as Edge;
                if (ed != null)
                {
                    object geo = null;
                    try { geo = ed.Geometry; } catch { geo = null; }
                    if (geo != null)
                    {
                        if (TryGetAxisFromAnyObject(geo, out axisWorldDir, out axisWorldPoint, out source))
                        {
                            source = "Edge.Geometry→" + source;
                            return true;
                        }
                    }
                }

                Inventor.Face f = obj as Inventor.Face;
                if (f != null)
                {
                    object geo = null;
                    try { geo = f.Geometry; } catch { geo = null; }
                    if (geo != null)
                    {
                        if (TryGetAxisFromAnyObject(geo, out axisWorldDir, out axisWorldPoint, out source))
                        {
                            source = "Face.Geometry→" + source;
                            return true;
                        }
                    }
                }
            }
            catch { }

            return false;
        }

        // =====================================================
        //  AxisWorldToOccLocal: axis_local = R^T * axis_world
        // =====================================================
        
    

        // =====================================================
        //  Matrix -> RPY (rad)  (roll X, pitch Y, yaw Z)
        // =====================================================

       

       
// ========================================================================
//  BLOQUE 2/4 (continuación dentro de UrdfExporter)
//  - Tessellate (CalculateFacets)
//  - TransformVerticesToLocalFrame
//  - Color/Atlas helpers (prioridad de color)
//  - ExportGeometryAndTextures (Part/Assembly)
//  - UpdateJointTypesFromConstraints (constraints -> joints)
// ========================================================================

        // =====================================================
        //  TESSELLATE (usa CalculateFacets)
        // =====================================================
        private static bool TessellateBodiesToMeshArrays(
            IList<SurfaceBody> bodies,
            out double[] vertices,
            out int[] indices)
        {
            vertices = null;
            indices = null;

            if (bodies == null || bodies.Count == 0)
            {
                DebugLog("MESH", "TessellateBodiesToMeshArrays: bodies == null o Count == 0");
                return false;
            }

            List<double> vList = new List<double>();
            List<int> iList = new List<int>();
            int vertexOffset = 0;

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
            indices = iList.ToArray();

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
                int facetCount = 0;

                double[] vertexCoords = new double[] { };
                double[] normalVectors = new double[] { };
                int[] vertexIndices = new int[] { };

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
                    ", facetCount=" + facetCount.ToString(CultureInfo.InvariantCulture));

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
                    double vCm = vertexCoords[i];
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
                double vx = verticesWorld[i] - tx;
                double vy = verticesWorld[i + 1] - ty;
                double vz = verticesWorld[i + 2] - tz;

                double lx = r11 * vx + r21 * vy + r31 * vz;
                double ly = r12 * vx + r22 * vy + r32 * vz;
                double lz = r13 * vx + r23 * vy + r33 * vz;

                verticesLocal[i] = lx;
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
                DebugLog("MESH", "LogAssetInfo: " + ownerKind + "='" + ownerName + "' sin Asset (null).");
                return;
            }

            string appDisplayName = "(sin nombre)";
            try { appDisplayName = app.DisplayName; } catch { appDisplayName = "(error DisplayName)"; }

            int count = 0;
            try { count = app.Count; } catch { count = -1; }

            DebugLog("MESH",
                "LogAssetInfo: " + ownerKind +
                "='" + ownerName +
                "', Asset.DisplayName='" + appDisplayName +
                "', AssetValues: Count=" + count.ToString(CultureInfo.InvariantCulture));

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
                        DebugLog("MESH", "      [Error leyendo ColorAssetValue.Value]");
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
        // =====================================================
        private static bool TryGetColorFromNamedAssetValue(
            Asset app,
            string targetName,
            out double r,
            out double g,
            out double b)
        {
            r = 0.8; g = 0.8; b = 0.8;
            if (app == null || string.IsNullOrEmpty(targetName)) return false;

            try
            {
                foreach (AssetValue av in app)
                {
                    if (av == null) continue;

                    string avName = "";
                    try { avName = av.Name; } catch { avName = ""; }
                    if (string.IsNullOrEmpty(avName)) continue;

                    if (!string.Equals(avName, targetName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (av.ValueType != AssetValueTypeEnum.kAssetValueTypeColor)
                        continue;

                    ColorAssetValue cav = av as ColorAssetValue;
                    if (cav == null) continue;

                    Inventor.Color invCol = cav.Value as Inventor.Color;
                    if (invCol == null) continue;

                    r = invCol.Red / 255.0;
                    g = invCol.Green / 255.0;
                    b = invCol.Blue / 255.0;

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
                DebugLog("MESH", "TryGetColorFromNamedAssetValue: error buscando '" + targetName + "'.");
            }

            return false;
        }

        // =====================================================
        //  Helper central: PRIORIDAD de color dentro de un Asset
        // =====================================================

        private static bool TryGetColorFromAssetWithPriority(
            Asset app,
            string ownerKind,
            string ownerName,
            out double r,
            out double g,
            out double b)
        {
                r = 0.8; g = 0.8; b = 0.8;

                if (app == null)
                {
                        DebugLog("MESH",
                            "TryGetColorFromAssetWithPriority: " + ownerKind +
                            "='" + ownerName + "' sin Asset, usando gris 0.8.");
                        return false;
                }

                LogAssetInfo(ownerKind, ownerName, app);

                // 1) generic_diffuse_color
                if (TryGetColorFromNamedAssetValue(app, "generic_diffuse_color", out r, out g, out b))
                        return true;

                // 2) generic_diffuse
                if (TryGetColorFromNamedAssetValue(app, "generic_diffuse", out r, out g, out b))
                        return true;

                // 3) metallicpaint_base_color
                if (TryGetColorFromNamedAssetValue(app, "metallicpaint_base_color", out r, out g, out b))
                        return true;

                // 4) plasticvinyl_color
                if (TryGetColorFromNamedAssetValue(app, "plasticvinyl_color", out r, out g, out b))
                        return true;

                // 5) wallpaint_color
                if (TryGetColorFromNamedAssetValue(app, "wallpaint_color", out r, out g, out b))
                        return true;

                // (Extra robustez: si el helper falla por cualquier razón, intenta acceso directo)
                try
                {
                        AssetValue avDif = null;

                        try { avDif = app["generic_diffuse_color"]; } catch { avDif = null; }
                        if (avDif == null)
                        {
                                try { avDif = app["generic_diffuse"]; } catch { avDif = null; }
                        }

                        if (avDif != null && avDif.ValueType == AssetValueTypeEnum.kAssetValueTypeColor)
                        {
                                ColorAssetValue difCav = avDif as ColorAssetValue;
                                if (difCav != null)
                                {
                                        Inventor.Color invCol1 = difCav.Value as Inventor.Color;
                                        if (invCol1 != null)
                                        {
                                                r = invCol1.Red / 255.0;
                                                g = invCol1.Green / 255.0;
                                                b = invCol1.Blue / 255.0;
                                                return true;
                                        }
                                }
                        }
                }
                catch { }

                // 6) common_tint_color (solo si no es grisáceo)
                try
                {
                        double tr, tg, tb;

                        bool gotTint =
                            TryGetColorFromNamedAssetValue(app, "common_tint_color", out tr, out tg, out tb) ||
                            TryGetColorFromNamedAssetValue(app, "common_Tint_color", out tr, out tg, out tb);

                        if (gotTint)
                        {
                                bool isGrayish =
                                    Math.Abs(tr - tg) < 0.02 &&
                                    Math.Abs(tg - tb) < 0.02;

                                if (!isGrayish)
                                {
                                        r = tr; g = tg; b = tb;
                                        return true;
                                }
                        }
                }
                catch { }

                // 7) fallback: primer AssetValue COLOR con DisplayName == "Color"
                try
                {
                        foreach (AssetValue av in app)
                        {
                                if (av == null) continue;
                                if (av.ValueType != AssetValueTypeEnum.kAssetValueTypeColor) continue;

                                string dn = null;
                                try { dn = av.DisplayName; } catch { dn = null; }
                                if (dn == null) continue;
                                if (!string.Equals(dn, "Color", StringComparison.OrdinalIgnoreCase)) continue;

                                ColorAssetValue cav = av as ColorAssetValue;
                                if (cav == null) continue;

                                Inventor.Color invCol = cav.Value as Inventor.Color;
                                if (invCol == null) continue;

                                r = invCol.Red / 255.0;
                                g = invCol.Green / 255.0;
                                b = invCol.Blue / 255.0;
                                return true;
                        }
                }
                catch { }

                // 8) fallback final: primer AssetValue COLOR cualquiera
                try
                {
                        foreach (AssetValue av in app)
                        {
                                if (av == null) continue;
                                if (av.ValueType != AssetValueTypeEnum.kAssetValueTypeColor) continue;

                                ColorAssetValue cav = av as ColorAssetValue;
                                if (cav == null) continue;

                                Inventor.Color invCol = cav.Value as Inventor.Color;
                                if (invCol == null) continue;

                                r = invCol.Red / 255.0;
                                g = invCol.Green / 255.0;
                                b = invCol.Blue / 255.0;
                                return true;
                        }
                }
                catch { }

                return false;
        }


        // =====================================================
        //  COLOR (Body y Face) + Fallback a Occurrence.Appearance
        // =====================================================
        private static bool TryGetBodyColor(
            SurfaceBody body,
            string ownerNameForLog,
            Asset occAppearance,
            out double r,
            out double g,
            out double b)
        {
            r = 0.8; g = 0.8; b = 0.8;

            if (body == null) return false;

            string bodyName = "(sin nombre)";
            try { if (!string.IsNullOrEmpty(body.Name)) bodyName = body.Name; } catch { }

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = bodyName;

            // 1) Body.Appearance
            try
            {
                Asset appBody = null;
                try { appBody = body.Appearance; } catch { appBody = null; }

                if (appBody != null &&
                    TryGetColorFromAssetWithPriority(appBody, "Body", ownerNameForLog, out r, out g, out b))
                {
                    return true;
                }
            }
            catch { }

            // 2) Fallback: Occurrence.Appearance
            if (occAppearance != null)
            {
                if (TryGetColorFromAssetWithPriority(occAppearance, "Occurrence", ownerNameForLog, out r, out g, out b))
                    return true;
            }

            return false;
        }

        private static bool TryGetBodyColor(
            SurfaceBody body,
            out double r,
            out double g,
            out double b)
        {
            return TryGetBodyColor(
                body,
                (body != null && !string.IsNullOrEmpty(body.Name)) ? body.Name : "(body)",
                null,
                out r, out g, out b);
        }

        private static bool TryGetFaceColor(
            Inventor.Face face,
            SurfaceBody parentBody,
            string ownerNameForLog,
            Asset occAppearance,
            out double r,
            out double g,
            out double b)
        {
            r = 0.8; g = 0.8; b = 0.8;

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = "(Face)";

            // 1) Face.Appearance
            if (face != null)
            {
                try
                {
                    Asset app = null;
                    try { app = face.Appearance; } catch { app = null; }

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
                            return true;
                    }
                }
                catch { }
            }

            // 2) Fallback: Body/Occurrence
            if (parentBody != null)
            {
                if (TryGetBodyColor(parentBody, ownerNameForLog, occAppearance, out r, out g, out b))
                    return true;
            }
            else if (occAppearance != null)
            {
                if (TryGetColorFromAssetWithPriority(occAppearance, "Occurrence", ownerNameForLog, out r, out g, out b))
                    return true;
            }

            return false;
        }

        private static int ClampToByte(double v)
        {
            if (v < 0.0) return 0;
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
            using (Bitmap bmp = new Bitmap(size, size))
            {
                System.Drawing.Color col = System.Drawing.Color.FromArgb(
                    255,
                    ClampToByte(r * 255.0),
                    ClampToByte(g * 255.0),
                    ClampToByte(b * 255.0));

                for (int y = 0; y < size; y++)
                    for (int x = 0; x < size; x++)
                        bmp.SetPixel(x, y, col);

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
            int width = cellsX * cellSize;
            int height = cellsY * cellSize;

            using (Bitmap bmp = new Bitmap(width, height))
            {
                System.Drawing.Color col = System.Drawing.Color.FromArgb(
                    255,
                    ClampToByte(r * 255.0),
                    ClampToByte(g * 255.0),
                    ClampToByte(b * 255.0));

                for (int y = 0; y < height; y++)
                    for (int x = 0; x < width; x++)
                        bmp.SetPixel(x, y, col);

                bmp.Save(path, ImageFormat.Png);
            }
        }

        private static void WriteBodyFaceColorAtlasCore(
            SurfaceBody body,
            string ownerNameForLog,
            Asset occAppearance,
            string path,
            int cellSize)
        {
            if (body == null)
            {
                WriteSolidColorPng(path, 0.8, 0.8, 0.8, cellSize);
                return;
            }

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = body.Name ?? "(body)";

            double bodyR, bodyG, bodyB;
            if (!TryGetBodyColor(body, ownerNameForLog, occAppearance, out bodyR, out bodyG, out bodyB))
                bodyR = bodyG = bodyB = 0.8;

            Faces faces = null;
            try { faces = body.Faces; } catch { faces = null; }

            int faceCount = (faces != null) ? faces.Count : 0;

            if (faceCount <= 0)
            {
                WriteAtlasSingleColorPng(path, bodyR, bodyG, bodyB, 1, 1, cellSize);
                return;
            }

            int cellsX = (int)Math.Ceiling(Math.Sqrt((double)faceCount));
            if (cellsX < 1) cellsX = 1;
            int cellsY = (int)Math.Ceiling((double)faceCount / (double)cellsX);
            if (cellsY < 1) cellsY = 1;

            int width = cellsX * cellSize;
            int height = cellsY * cellSize;

            using (Bitmap bmp = new Bitmap(width, height))
            {
                using (Graphics gg = Graphics.FromImage(bmp))
                {
                    System.Drawing.Color bgCol = System.Drawing.Color.FromArgb(
                        255,
                        ClampToByte(bodyR * 255.0),
                        ClampToByte(bodyG * 255.0),
                        ClampToByte(bodyB * 255.0));
                    gg.Clear(bgCol);
                }

                for (int fi = 0; fi < faceCount; fi++)
                {
                    Inventor.Face f = null;
                    try { f = faces[fi + 1]; } catch { f = null; }

                    double fr, fg, fb;
                    if (!TryGetFaceColor(f, body, ownerNameForLog, occAppearance, out fr, out fg, out fb))
                    {
                        fr = bodyR; fg = bodyG; fb = bodyB;
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

                    for (int y = startY; y < startY + cellSize && y < height; y++)
                        for (int x = startX; x < startX + cellSize && x < width; x++)
                            bmp.SetPixel(x, y, faceCol);
                }

                bmp.Save(path, ImageFormat.Png);
            }
        }

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

        private static void WriteBodyFaceColorAtlas(
            SurfaceBody body,
            string ownerNameForLog,
            Asset occAppearance,
            string path,
            int cellSize)
        {
            WriteBodyFaceColorAtlasCore(body, ownerNameForLog, occAppearance, path, cellSize);
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

            for (int i = 0; i < bodies.Count; i++)
            {
                SurfaceBody body = bodies[i];
                if (body == null) continue;

                string bodyName = "(null)";
                try { if (!string.IsNullOrEmpty(body.Name)) bodyName = body.Name; } catch { }

                string linkName = "root_body_" +
                                  i.ToString(CultureInfo.InvariantCulture) + "_" +
                                  MakeSafeName(bodyName);

                UrdfLink link = FindLinkByName(robot, linkName);
                if (link == null) continue;

                double[] vertices;
                int[] indices;

                List<SurfaceBody> oneBodyList = new List<SurfaceBody>();
                oneBodyList.Add(body);

                if (!TessellateBodiesToMeshArrays(oneBodyList, out vertices, out indices))
                    continue;

                string daeName = linkName + ".dae";
                string daePath = IOPath.Combine(meshesDir, daeName);

                // (WriteColladaFile está en BLOQUE 3/4)
                WriteColladaFile(daePath, linkName, vertices, indices);
                link.MeshFile = "meshes/" + daeName;

                // ====== COLOR / TEXTURA ======
                double r, g, b;
                if (!TryGetBodyColor(body, linkName, null, out r, out g, out b))
                    r = g = b = 0.8;

                string pngPath = IOPath.Combine(meshesDir, linkName + ".png");

                if (_meshQualityMode == "low")
                    WriteSolidColorPng(pngPath, r, g, b, 32);
                else
                    WriteBodyFaceColorAtlas(body, linkName, null, pngPath, 32);

                // ====== INERCIA ======
                try
                {
                    MassProperties mp = partDef.MassProperties;
                    // (FillLinkInertialFromMassProperties está en BLOQUE 3/4)
                    FillLinkInertialFromMassProperties(link, mp);
                }
                catch { }
            }
        }

        private static void ExportAssemblyGeometryToDae(
            AssemblyDocument asmDoc,
            RobotModel robot,
            string meshesDir)
        {
            if (asmDoc == null || robot == null) return;

            // 🔧 Ajuste JOINTS desde constraints (robusto)
            UpdateJointTypesFromConstraints(asmDoc, robot);

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs = asmDef.Occurrences;
            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            int occIndex = 0;

            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    if (occ == null) { occIndex++; continue; }
                    if (occ.Suppressed) { occIndex++; continue; }
                    if (!occ.Visible) { occIndex++; continue; }

                    string rawName = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    string baseLinkName = "link_" +
                                          occIndex.ToString(CultureInfo.InvariantCulture) +
                                          "_" + safeName;

                    // Apariencia a nivel de OCCURRENCE
                    Asset occAppearance = null;
                    try { occAppearance = occ.Appearance; } catch { occAppearance = null; }

                    List<SurfaceBody> bodies = new List<SurfaceBody>();
                    CollectSurfaceBodiesFromOccurrence(occ, bodies);

                    if (bodies.Count == 0) { occIndex++; continue; }

                    Matrix m = occ.Transformation;

                    for (int i = 0; i < bodies.Count; i++)
                    {
                        SurfaceBody body = bodies[i];
                        if (body == null) continue;

                        string suffix = (i == 0) ? "" : "_b" + i.ToString(CultureInfo.InvariantCulture);
                        string linkName = baseLinkName + suffix;

                        UrdfLink link = FindLinkByName(robot, linkName);
                        if (link == null) continue;

                        double[] verticesWorld;
                        int[] indices;

                        List<SurfaceBody> oneBodyList = new List<SurfaceBody>();
                        oneBodyList.Add(body);

                        if (!TessellateBodiesToMeshArrays(oneBodyList, out verticesWorld, out indices))
                            continue;

                        double[] verticesLocal;
                        TransformVerticesToLocalFrame(verticesWorld, m, out verticesLocal);

                        string daeName = linkName + ".dae";
                        string daePath = IOPath.Combine(meshesDir, daeName);

                        // (WriteColladaFile está en BLOQUE 3/4)
                        WriteColladaFile(daePath, linkName, verticesLocal, indices);
                        link.MeshFile = "meshes/" + daeName;

                        // ====== COLOR / TEXTURA ======
                        double r, g, b;
                        if (!TryGetBodyColor(body, rawName, occAppearance, out r, out g, out b))
                            r = g = b = 0.8;

                        string pngPath = IOPath.Combine(meshesDir, linkName + ".png");

                        if (_meshQualityMode == "low")
                            WriteSolidColorPng(pngPath, r, g, b, 32);
                        else
                            WriteBodyFaceColorAtlas(body, rawName, occAppearance, pngPath, 32);

                        // ====== INERCIA ======
                        try
                        {
                            PartComponentDefinition partDef = occ.Definition as PartComponentDefinition;
                            if (partDef != null)
                            {
                                MassProperties mp = partDef.MassProperties;
                                // (FillLinkInertialFromMassProperties está en BLOQUE 3/4)
                                FillLinkInertialFromMassProperties(link, mp);
                            }
                        }
                        catch { }
                    }

                    occIndex++;
                }
                catch
                {
                    occIndex++;
                }
            }
        }

        // -------------------------------------------------
        //  Buscar link por nombre
        // -------------------------------------------------
        private static UrdfLink FindLinkByName(RobotModel robot, string name)
        {
            if (robot == null || robot.Links == null) return null;

            foreach (UrdfLink link in robot.Links)
                if (link != null && link.Name == name)
                    return link;

            return null;
        }

        // -------------------------------------------------
        //  Buscar JOINT por ChildLink (para mapear constraints)
        // -------------------------------------------------
       
        // -------------------------------------------------
        //  Resolver occurrences de constraints a una leaf occurrence mapeada
        // -------------------------------------------------
        

        // -------------------------------------------------
        //  Mapear AssemblyConstraints → tipos de JOINT URDF (ROBUSTO)
        //  (incluye: parent real, origin relativo, axis real -> axis local)
        // -------------------------------------------------
        

        // =====================================================
        //  RELATIVE origin: parent→child (m) + RPY(rad)
        // =====================================================
        

        // =====================================================
        //  EXTRAER AXIS WORLD DESDE CONSTRAINT
        // =====================================================
        private static bool TryExtractAxisWorldFromConstraint(
            AssemblyConstraint ac,
            out double[] axisWorld,
            out string axisSource)
        {
            axisWorld = null;
            axisSource = null;
            if (ac == null) return false;

            object g1 = null, g2 = null, e1 = null, e2 = null;

            try
            {
                if (ac is MateConstraint)
                {
                    MateConstraint mc = (MateConstraint)ac;
                    try { g1 = mc.GeometryOne; } catch { }
                    try { g2 = mc.GeometryTwo; } catch { }
                    try { e1 = mc.EntityOne; } catch { }
                    try { e2 = mc.EntityTwo; } catch { }
                }
                else if (ac is InsertConstraint)
                {
                    InsertConstraint ic = (InsertConstraint)ac;
                    try { g1 = ic.GeometryOne; } catch { }
                    try { g2 = ic.GeometryTwo; } catch { }
                    try { e1 = ic.EntityOne; } catch { }
                    try { e2 = ic.EntityTwo; } catch { }
                }
                else if (ac is AngleConstraint)
                {
                    AngleConstraint ang = (AngleConstraint)ac;
                    try { g1 = ang.GeometryOne; } catch { }
                    try { g2 = ang.GeometryTwo; } catch { }
                    try { e1 = ang.EntityOne; } catch { }
                    try { e2 = ang.EntityTwo; } catch { }
                }
                else if (ac is TransitionalConstraint)
                {
                    TransitionalConstraint tc = (TransitionalConstraint)ac;
                    try { g1 = tc.GeometryOne; } catch { }
                    try { g2 = tc.GeometryTwo; } catch { }
                    try { e1 = tc.EntityOne; } catch { }
                    try { e2 = tc.EntityTwo; } catch { }
                }
                else
                {
                    try { e1 = ac.EntityOne; } catch { }
                    try { e2 = ac.EntityTwo; } catch { }
                    try { g1 = ac.GeometryOne; } catch { }
                    try { g2 = ac.GeometryTwo; } catch { }
                }
            }
            catch { }

            if (TryGetAxisFromAnyObject(g1, out axisWorld, out axisSource)) return true;
            if (TryGetAxisFromAnyObject(g2, out axisWorld, out axisSource)) return true;
            if (TryGetAxisFromAnyObject(e1, out axisWorld, out axisSource)) return true;
            if (TryGetAxisFromAnyObject(e2, out axisWorld, out axisSource)) return true;

            return false;
        }

        private static bool TryGetAxisFromAnyObject(
            object obj,
            out double[] axisWorld,
            out string source)
        {
            axisWorld = null;
            source = null;
            if (obj == null) return false;

            try
            {
                Cylinder cyl = obj as Cylinder;
                if (cyl != null)
                {
                    UnitVector uv = cyl.AxisVector;
                    axisWorld = Normalize3(new double[] { uv.X, uv.Y, uv.Z });
                    source = "Cylinder.AxisVector";
                    return axisWorld != null;
                }

                Cone cone = obj as Cone;
                if (cone != null)
                {
                    UnitVector uv = cone.AxisVector;
                    axisWorld = Normalize3(new double[] { uv.X, uv.Y, uv.Z });
                    source = "Cone.AxisVector";
                    return axisWorld != null;
                }

                Line line = obj as Line;
                if (line != null)
                {
                    UnitVector uv = line.Direction;
                    axisWorld = Normalize3(new double[] { uv.X, uv.Y, uv.Z });
                    source = "Line.Direction";
                    return axisWorld != null;
                }

                LineSegment seg = obj as LineSegment;
                if (seg != null)
                {
                    UnitVector uv = seg.Direction;
                    axisWorld = Normalize3(new double[] { uv.X, uv.Y, uv.Z });
                    source = "LineSegment.Direction";
                    return axisWorld != null;
                }

                WorkAxis wa = obj as WorkAxis;
                if (wa != null)
                {
                    Line wl = null;
                    try { wl = wa.Line; } catch { wl = null; }
                    if (wl != null)
                    {
                        UnitVector uv = wl.Direction;
                        axisWorld = Normalize3(new double[] { uv.X, uv.Y, uv.Z });
                        source = "WorkAxis.Line.Direction";
                        return axisWorld != null;
                    }
                }

                Edge ed = obj as Edge;
                if (ed != null)
                {
                    object geo = null;
                    try { geo = ed.Geometry; } catch { geo = null; }
                    if (geo != null)
                    {
                        if (TryGetAxisFromAnyObject(geo, out axisWorld, out source))
                        {
                            source = "Edge.Geometry→" + source;
                            return true;
                        }
                    }
                }

                Inventor.Face f = obj as Inventor.Face;
                if (f != null)
                {
                    object geo = null;
                    try { geo = f.Geometry; } catch { geo = null; }
                    if (geo != null)
                    {
                        if (TryGetAxisFromAnyObject(geo, out axisWorld, out source))
                        {
                            source = "Face.Geometry→" + source;
                            return true;
                        }
                    }
                }
            }
            catch { }

            return false;
        }

        // =====================================================
        //  AxisWorldToOccLocal:
        //  axis_local = R^T * axis_world
        // =====================================================
        private static double[] AxisWorldToOccLocal(double[] axisWorld, Matrix occMatrix)
        {
            if (axisWorld == null || axisWorld.Length != 3) return null;

            double[] a = Normalize3(new double[] { axisWorld[0], axisWorld[1], axisWorld[2] });
            if (a == null) return null;

            if (occMatrix == null) return a;

            double r11 = occMatrix.Cell[1, 1];
            double r12 = occMatrix.Cell[1, 2];
            double r13 = occMatrix.Cell[1, 3];

            double r21 = occMatrix.Cell[2, 1];
            double r22 = occMatrix.Cell[2, 2];
            double r23 = occMatrix.Cell[2, 3];

            double r31 = occMatrix.Cell[3, 1];
            double r32 = occMatrix.Cell[3, 2];
            double r33 = occMatrix.Cell[3, 3];

            double lx = r11 * a[0] + r21 * a[1] + r31 * a[2];
            double ly = r12 * a[0] + r22 * a[1] + r32 * a[2];
            double lz = r13 * a[0] + r23 * a[1] + r33 * a[2];

            return Normalize3(new double[] { lx, ly, lz });
        }

        private static double[] Normalize3(double[] v)
        {
            if (v == null || v.Length != 3) return null;
            double x = v[0], y = v[1], z = v[2];
            double n = Math.Sqrt(x * x + y * y + z * z);
            if (n < 1e-12) return null;
            return new double[] { x / n, y / n, z / n };
        }

        // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        //  Continúa en BLOQUE 3/4:
        //   - WriteColladaFile + Normals + UVs + XmlEscape/FormatF
        //   - FillLinkInertialFromMassProperties
        // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


















































// ========================================================================
//  BLOQUE 3/4 (continuación dentro de UrdfExporter)
//  - WriteColladaFile (DAE) + normals + uvs + material/texture
//  - Helpers: XmlEscape, F, EnsureDirectory (si no estaba), ComputeNormals
//  - FillLinkInertialFromMassProperties (MassProperties -> URDF inertia)
// ========================================================================

        // =====================================================
        //  COLLADA (.dae) writer (mínimo + texture .png)
        // =====================================================
        private static void WriteColladaFile(
            string daePath,
            string meshName,
            double[] vertices,
            int[] indices)
        {
            try
            {
                if (string.IsNullOrEmpty(daePath) || vertices == null || vertices.Length == 0 ||
                    indices == null || indices.Length == 0)
                {
                    DebugLog("ERR", "WriteColladaFile: parámetros inválidos.");
                    return;
                }

                int vCount = vertices.Length / 3;
                int triCount = indices.Length / 3;

                if (vCount <= 0 || triCount <= 0)
                {
                    DebugLog("ERR", "WriteColladaFile: vCount<=0 o triCount<=0.");
                    return;
                }

                // Normal por vértice (acumulada)
                double[] normals = ComputeVertexNormals(vertices, indices);

                // UVs simples (toda la malla apunta al centro de la textura)
                // Si usas atlas por cara, esto mostrará SOLO una celda (placeholder).
                double[] uvs = new double[vCount * 2];
                for (int i = 0; i < vCount; i++)
                {
                    uvs[i * 2 + 0] = 0.5;
                    uvs[i * 2 + 1] = 0.5;
                }

                string daeDir = IOPath.GetDirectoryName(daePath);
                if (!string.IsNullOrEmpty(daeDir)) EnsureDirectory(daeDir);

                // PNG esperado: mismo basename
                string pngFile = meshName + ".png";

                string safeMesh = MakeSafeName(meshName);
                string geoId = safeMesh + "_geo";
                string posId = safeMesh + "_pos";
                string norId = safeMesh + "_nor";
                string uvId = safeMesh + "_uv";
                string vtxId = safeMesh + "_vtx";

                string imgId = safeMesh + "_img";
                string effId = safeMesh + "_eff";
                string matId = safeMesh + "_mat";
                string sceneId = "Scene";
                string nodeId = safeMesh + "_node";

                StringBuilder sb = new StringBuilder(1024 * 64);

                sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
                sb.Append("<COLLADA xmlns=\"http://www.collada.org/2005/11/COLLADASchema\" version=\"1.4.1\">\n");

                // Asset
                sb.Append("  <asset>\n");
                sb.Append("    <contributor><authoring_tool>Inventor URDFConverter</authoring_tool></contributor>\n");
                sb.Append("    <unit name=\"meter\" meter=\"1\"/>\n");
                sb.Append("    <up_axis>Z_UP</up_axis>\n");
                sb.Append("  </asset>\n");

                // Images
                sb.Append("  <library_images>\n");
                sb.Append("    <image id=\"").Append(XmlEscape(imgId)).Append("\" name=\"").Append(XmlEscape(imgId)).Append("\">\n");
                sb.Append("      <init_from>").Append(XmlEscape(pngFile)).Append("</init_from>\n");
                sb.Append("    </image>\n");
                sb.Append("  </library_images>\n");

                // Effects (Lambert + texture)
                sb.Append("  <library_effects>\n");
                sb.Append("    <effect id=\"").Append(XmlEscape(effId)).Append("\" name=\"").Append(XmlEscape(effId)).Append("\">\n");
                sb.Append("      <profile_COMMON>\n");
                sb.Append("        <newparam sid=\"surface\">\n");
                sb.Append("          <surface type=\"2D\">\n");
                sb.Append("            <init_from>").Append(XmlEscape(imgId)).Append("</init_from>\n");
                sb.Append("          </surface>\n");
                sb.Append("        </newparam>\n");
                sb.Append("        <newparam sid=\"sampler\">\n");
                sb.Append("          <sampler2D>\n");
                sb.Append("            <source>surface</source>\n");
                sb.Append("          </sampler2D>\n");
                sb.Append("        </newparam>\n");
                sb.Append("        <technique sid=\"common\">\n");
                sb.Append("          <lambert>\n");
                sb.Append("            <diffuse>\n");
                sb.Append("              <texture texture=\"sampler\" texcoord=\"UVSET0\"/>\n");
                sb.Append("            </diffuse>\n");
                sb.Append("          </lambert>\n");
                sb.Append("        </technique>\n");
                sb.Append("      </profile_COMMON>\n");
                sb.Append("    </effect>\n");
                sb.Append("  </library_effects>\n");

                // Materials
                sb.Append("  <library_materials>\n");
                sb.Append("    <material id=\"").Append(XmlEscape(matId)).Append("\" name=\"").Append(XmlEscape(matId)).Append("\">\n");
                sb.Append("      <instance_effect url=\"#").Append(XmlEscape(effId)).Append("\"/>\n");
                sb.Append("    </material>\n");
                sb.Append("  </library_materials>\n");

                // Geometries
                sb.Append("  <library_geometries>\n");
                sb.Append("    <geometry id=\"").Append(XmlEscape(geoId)).Append("\" name=\"").Append(XmlEscape(geoId)).Append("\">\n");
                sb.Append("      <mesh>\n");

                // Positions source
                sb.Append("        <source id=\"").Append(XmlEscape(posId)).Append("\">\n");
                sb.Append("          <float_array id=\"").Append(XmlEscape(posId)).Append("_arr\" count=\"")
                  .Append((vCount * 3).ToString(CultureInfo.InvariantCulture)).Append("\">");
                for (int i = 0; i < vertices.Length; i++)
                {
                    sb.Append(F(vertices[i])).Append(i + 1 < vertices.Length ? " " : "");
                }
                sb.Append("</float_array>\n");
                sb.Append("          <technique_common>\n");
                sb.Append("            <accessor source=\"#").Append(XmlEscape(posId)).Append("_arr\" count=\"")
                  .Append(vCount.ToString(CultureInfo.InvariantCulture))
                  .Append("\" stride=\"3\">\n");
                sb.Append("              <param name=\"X\" type=\"float\"/>\n");
                sb.Append("              <param name=\"Y\" type=\"float\"/>\n");
                sb.Append("              <param name=\"Z\" type=\"float\"/>\n");
                sb.Append("            </accessor>\n");
                sb.Append("          </technique_common>\n");
                sb.Append("        </source>\n");

                // Normals source
                sb.Append("        <source id=\"").Append(XmlEscape(norId)).Append("\">\n");
                sb.Append("          <float_array id=\"").Append(XmlEscape(norId)).Append("_arr\" count=\"")
                  .Append((vCount * 3).ToString(CultureInfo.InvariantCulture)).Append("\">");
                for (int i = 0; i < normals.Length; i++)
                {
                    sb.Append(F(normals[i])).Append(i + 1 < normals.Length ? " " : "");
                }
                sb.Append("</float_array>\n");
                sb.Append("          <technique_common>\n");
                sb.Append("            <accessor source=\"#").Append(XmlEscape(norId)).Append("_arr\" count=\"")
                  .Append(vCount.ToString(CultureInfo.InvariantCulture))
                  .Append("\" stride=\"3\">\n");
                sb.Append("              <param name=\"X\" type=\"float\"/>\n");
                sb.Append("              <param name=\"Y\" type=\"float\"/>\n");
                sb.Append("              <param name=\"Z\" type=\"float\"/>\n");
                sb.Append("            </accessor>\n");
                sb.Append("          </technique_common>\n");
                sb.Append("        </source>\n");

                // UV source
                sb.Append("        <source id=\"").Append(XmlEscape(uvId)).Append("\">\n");
                sb.Append("          <float_array id=\"").Append(XmlEscape(uvId)).Append("_arr\" count=\"")
                  .Append((vCount * 2).ToString(CultureInfo.InvariantCulture)).Append("\">");
                for (int i = 0; i < uvs.Length; i++)
                {
                    sb.Append(F(uvs[i])).Append(i + 1 < uvs.Length ? " " : "");
                }
                sb.Append("</float_array>\n");
                sb.Append("          <technique_common>\n");
                sb.Append("            <accessor source=\"#").Append(XmlEscape(uvId)).Append("_arr\" count=\"")
                  .Append(vCount.ToString(CultureInfo.InvariantCulture))
                  .Append("\" stride=\"2\">\n");
                sb.Append("              <param name=\"S\" type=\"float\"/>\n");
                sb.Append("              <param name=\"T\" type=\"float\"/>\n");
                sb.Append("            </accessor>\n");
                sb.Append("          </technique_common>\n");
                sb.Append("        </source>\n");

                // Vertices
                sb.Append("        <vertices id=\"").Append(XmlEscape(vtxId)).Append("\">\n");
                sb.Append("          <input semantic=\"POSITION\" source=\"#").Append(XmlEscape(posId)).Append("\"/>\n");
                sb.Append("        </vertices>\n");

                // Triangles (VERTEX + NORMAL + TEXCOORD)
                sb.Append("        <triangles count=\"").Append(triCount.ToString(CultureInfo.InvariantCulture))
                  .Append("\" material=\"").Append(XmlEscape(matId)).Append("\">\n");
                sb.Append("          <input semantic=\"VERTEX\" source=\"#").Append(XmlEscape(vtxId)).Append("\" offset=\"0\"/>\n");
                sb.Append("          <input semantic=\"NORMAL\" source=\"#").Append(XmlEscape(norId)).Append("\" offset=\"1\"/>\n");
                sb.Append("          <input semantic=\"TEXCOORD\" source=\"#").Append(XmlEscape(uvId)).Append("\" offset=\"2\" set=\"0\"/>\n");
                sb.Append("          <p>");

                // Para cada índice de vértice vi, repetimos vi para VERTEX/NORMAL/TEXCOORD
                for (int k = 0; k < indices.Length; k++)
                {
                    int vi = indices[k];
                    if (vi < 0) vi = 0;
                    if (vi >= vCount) vi = vCount - 1;

                    sb.Append(vi.ToString(CultureInfo.InvariantCulture)).Append(" ");
                    sb.Append(vi.ToString(CultureInfo.InvariantCulture)).Append(" ");
                    sb.Append(vi.ToString(CultureInfo.InvariantCulture));

                    if (k + 1 < indices.Length) sb.Append(" ");
                }

                sb.Append("</p>\n");
                sb.Append("        </triangles>\n");

                sb.Append("      </mesh>\n");
                sb.Append("    </geometry>\n");
                sb.Append("  </library_geometries>\n");

                // Visual scenes
                sb.Append("  <library_visual_scenes>\n");
                sb.Append("    <visual_scene id=\"").Append(XmlEscape(sceneId)).Append("\" name=\"").Append(XmlEscape(sceneId)).Append("\">\n");
                sb.Append("      <node id=\"").Append(XmlEscape(nodeId)).Append("\" name=\"").Append(XmlEscape(nodeId)).Append("\">\n");
                sb.Append("        <instance_geometry url=\"#").Append(XmlEscape(geoId)).Append("\">\n");
                sb.Append("          <bind_material>\n");
                sb.Append("            <technique_common>\n");
                sb.Append("              <instance_material symbol=\"").Append(XmlEscape(matId)).Append("\" target=\"#").Append(XmlEscape(matId)).Append("\">\n");
                sb.Append("                <bind_vertex_input semantic=\"UVSET0\" input_semantic=\"TEXCOORD\" input_set=\"0\"/>\n");
                sb.Append("              </instance_material>\n");
                sb.Append("            </technique_common>\n");
                sb.Append("          </bind_material>\n");
                sb.Append("        </instance_geometry>\n");
                sb.Append("      </node>\n");
                sb.Append("    </visual_scene>\n");
                sb.Append("  </library_visual_scenes>\n");

                // Scene
                sb.Append("  <scene>\n");
                sb.Append("    <instance_visual_scene url=\"#").Append(XmlEscape(sceneId)).Append("\"/>\n");
                sb.Append("  </scene>\n");

                sb.Append("</COLLADA>\n");

                IOFile.WriteAllText(daePath, sb.ToString(), Encoding.UTF8);

                DebugLog("MESH",
                    "WriteColladaFile: OK '" + daePath +
                    "', v=" + vCount.ToString(CultureInfo.InvariantCulture) +
                    ", tris=" + triCount.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                DebugLog("ERR", "WriteColladaFile: " + ex.Message);
            }
        }

        private static double[] ComputeVertexNormals(double[] vertices, int[] indices)
        {
            int vCount = vertices.Length / 3;
            double[] n = new double[vCount * 3];

            for (int t = 0; t + 2 < indices.Length; t += 3)
            {
                int i0 = indices[t + 0];
                int i1 = indices[t + 1];
                int i2 = indices[t + 2];

                if (i0 < 0 || i1 < 0 || i2 < 0) continue;
                if (i0 >= vCount || i1 >= vCount || i2 >= vCount) continue;

                double x0 = vertices[i0 * 3 + 0], y0 = vertices[i0 * 3 + 1], z0 = vertices[i0 * 3 + 2];
                double x1 = vertices[i1 * 3 + 0], y1 = vertices[i1 * 3 + 1], z1 = vertices[i1 * 3 + 2];
                double x2 = vertices[i2 * 3 + 0], y2 = vertices[i2 * 3 + 1], z2 = vertices[i2 * 3 + 2];

                double ax = x1 - x0, ay = y1 - y0, az = z1 - z0;
                double bx = x2 - x0, by = y2 - y0, bz = z2 - z0;

                // cross(a,b)
                double nx = ay * bz - az * by;
                double ny = az * bx - ax * bz;
                double nz = ax * by - ay * bx;

                n[i0 * 3 + 0] += nx; n[i0 * 3 + 1] += ny; n[i0 * 3 + 2] += nz;
                n[i1 * 3 + 0] += nx; n[i1 * 3 + 1] += ny; n[i1 * 3 + 2] += nz;
                n[i2 * 3 + 0] += nx; n[i2 * 3 + 1] += ny; n[i2 * 3 + 2] += nz;
            }

            // Normalize
            for (int i = 0; i < vCount; i++)
            {
                double x = n[i * 3 + 0], y = n[i * 3 + 1], z = n[i * 3 + 2];
                double len = Math.Sqrt(x * x + y * y + z * z);
                if (len < 1e-12)
                {
                    n[i * 3 + 0] = 0;
                    n[i * 3 + 1] = 0;
                    n[i * 3 + 2] = 1;
                }
                else
                {
                    n[i * 3 + 0] = x / len;
                    n[i * 3 + 1] = y / len;
                    n[i * 3 + 2] = z / len;
                }
            }

            return n;
        }

        // =====================================================
        //  URDF Inertial fill desde MassProperties
        // =====================================================
        private static void FillLinkInertialFromMassProperties(UrdfLink link, MassProperties mp)
        {
            if (link == null || mp == null) return;

            double cmToM = 0.01;     // Inventor length suele estar en cm
            double cm2ToM2 = 1e-4;   // kg*cm^2 -> kg*m^2

            // Mass
            double mass = 0.001;
            try
            {
                mass = mp.Mass;
                if (mass <= 0.0) mass = 0.001;
            }
            catch { mass = 0.001; }

            // COM
            double cx = 0, cy = 0, cz = 0;
            try
            {
                object comObj = null;
                try { comObj = mp.CenterOfMass; } catch { comObj = null; }

                InvPoint com = comObj as InvPoint;
                if (com != null)
                {
                    cx = com.X * cmToM;
                    cy = com.Y * cmToM;
                    cz = com.Z * cmToM;
                }
            }
            catch { }

            // Inertia defaults (si no se puede leer)
            double ixx = 1e-6, iyy = 1e-6, izz = 1e-6, ixy = 0, ixz = 0, iyz = 0;
            bool got = false;

            // 1) Intentar InertiaTensor por reflection (property o method)
            try
            {
                object itObj = TryComGet(mp, "InertiaTensor");
                Array itArr = itObj as Array;
                if (itArr != null && itArr.Length >= 9)
                {
                    // row-major:
                    // [0]=xx [1]=xy [2]=xz
                    // [3]=yx [4]=yy [5]=yz
                    // [6]=zx [7]=zy [8]=zz
                    double t00 = ToDoubleSafe(itArr.GetValue(0));
                    double t01 = ToDoubleSafe(itArr.GetValue(1));
                    double t02 = ToDoubleSafe(itArr.GetValue(2));
                    double t10 = ToDoubleSafe(itArr.GetValue(3));
                    double t11 = ToDoubleSafe(itArr.GetValue(4));
                    double t12 = ToDoubleSafe(itArr.GetValue(5));
                    double t20 = ToDoubleSafe(itArr.GetValue(6));
                    double t21 = ToDoubleSafe(itArr.GetValue(7));
                    double t22 = ToDoubleSafe(itArr.GetValue(8));

                    ixx = t00 * cm2ToM2;
                    iyy = t11 * cm2ToM2;
                    izz = t22 * cm2ToM2;

                    ixy = 0.5 * (t01 + t10) * cm2ToM2;
                    ixz = 0.5 * (t02 + t20) * cm2ToM2;
                    iyz = 0.5 * (t12 + t21) * cm2ToM2;

                    got = true;
                }
            }
            catch { got = false; }

            // 2) Fallback: PrincipalMomentsOfInertia por reflection (property o method)
            if (!got)
            {
                try
                {
                    object pmObj = TryComGet(mp, "PrincipalMomentsOfInertia");
                    Array pmArr = pmObj as Array;
                    if (pmArr != null && pmArr.Length >= 3)
                    {
                        ixx = ToDoubleSafe(pmArr.GetValue(0)) * cm2ToM2;
                        iyy = ToDoubleSafe(pmArr.GetValue(1)) * cm2ToM2;
                        izz = ToDoubleSafe(pmArr.GetValue(2)) * cm2ToM2;

                        ixy = ixz = iyz = 0.0;
                        got = true;
                    }
                }
                catch { }
            }

            // Guardar en link
            link.InertialMass = mass;
            link.InertialOriginXYZ = new double[] { cx, cy, cz };
            link.InertialOriginRPY = new double[] { 0.0, 0.0, 0.0 };

            link.Ixx = ixx;
            link.Iyy = iyy;
            link.Izz = izz;
            link.Ixy = ixy;
            link.Ixz = ixz;
            link.Iyz = iyz;
        }

        private static object TryComGet(object comObj, string member)
        {
            if (comObj == null || string.IsNullOrEmpty(member)) return null;

            try
            {
                return comObj.GetType().InvokeMember(
                    member,
                    BindingFlags.GetProperty,
                    null,
                    comObj,
                    null,
                    CultureInfo.InvariantCulture
                );
            }
            catch { }

            try
            {
                return comObj.GetType().InvokeMember(
                    member,
                    BindingFlags.InvokeMethod,
                    null,
                    comObj,
                    null,
                    CultureInfo.InvariantCulture
                );
            }
            catch { }

            return null;
        }

        private static double ToDoubleSafe(object o)
        {
            try
            {
                if (o == null) return 0.0;
                if (o is double) return (double)o;
                if (o is float) return (double)(float)o;
                if (o is int) return (double)(int)o;
                if (o is long) return (double)(long)o;
                return Convert.ToDouble(o, CultureInfo.InvariantCulture);
            }
            catch { return 0.0; }
        }


        // =====================================================
        //  Helpers XML / float formatting
        // =====================================================
        private static string XmlEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\"", "&quot;")
                    .Replace("'", "&apos;");
        }

        private static string F(double v)
        {
            return v.ToString("0.########", CultureInfo.InvariantCulture);
        }

        private static bool EnsureDirectory(string dir)
        {
            try
            {
                if (string.IsNullOrEmpty(dir)) return false;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return true;
            }
            catch
            {
                return false;
            }
        }


        // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        //  Continúa en BLOQUE 4/4:
        //   - WriteUrdfFile (robot.urdf)
        //   - Serialización Links/Joints (origin, axis, inertial)
        //   - Ensure every link has parent joint a base_link
        //   - Export entrypoints (botones) y “modo low/high”
        // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
































// ========================================================================
//  BLOQUE 4/4 (final)
//  - WriteUrdfFile (+ validación/auto-root joints)
//  - Helpers: FormatXYZ, EnsureBaseLink, EnsureEveryLinkHasParentJoint
//  - MatrixToRPY (si faltaba en tus bloques previos)
//  - Modelos: RobotModel, UrdfLink, UrdfJoint
// ========================================================================

        // =====================================================
        //  URDF writer (robot.urdf)
        // =====================================================
        private static void WriteUrdfFile(RobotModel robot, string urdfPath)
        {
            if (robot == null || string.IsNullOrEmpty(urdfPath)) return;

            try
            {
                EnsureBaseLink(robot);
                EnsureEveryLinkHasParentJoint(robot);

                // Validar joints: parent/child existentes
                List<UrdfJoint> validJoints = new List<UrdfJoint>();
                foreach (UrdfJoint j in robot.Joints)
                {
                    if (j == null) continue;
                    if (string.IsNullOrEmpty(j.Name)) continue;
                    if (string.IsNullOrEmpty(j.ChildLink)) continue;

                    UrdfLink child = FindLinkByName(robot, j.ChildLink);
                    if (child == null) continue;

                    if (string.IsNullOrEmpty(j.ParentLink))
                        j.ParentLink = "base_link";

                    UrdfLink parent = FindLinkByName(robot, j.ParentLink);
                    if (parent == null)
                    {
                        j.ParentLink = "base_link";
                        parent = FindLinkByName(robot, "base_link");
                        if (parent == null) continue;
                    }

                    // evitar joint parent==child
                    if (string.Equals(j.ParentLink, j.ChildLink, StringComparison.OrdinalIgnoreCase))
                        continue;

                    validJoints.Add(j);
                }

                StringBuilder sb = new StringBuilder(1024 * 256);

                sb.AppendLine("<?xml version=\"1.0\"?>");
                sb.AppendLine("<robot name=\"" + XmlEscape(robot.Name ?? "robot") + "\">");

                // ----------------
                // LINKS
                // ----------------
                foreach (UrdfLink link in robot.Links)
                {
                    if (link == null || string.IsNullOrEmpty(link.Name))
                        continue;

                    sb.AppendLine("  <link name=\"" + XmlEscape(link.Name) + "\">");

                    // inertial (siempre, con defaults)
                    double mass = (link.InertialMass > 1e-12) ? link.InertialMass : 0.01;
                    double[] com = (link.InertialOriginXYZ != null && link.InertialOriginXYZ.Length >= 3)
                        ? link.InertialOriginXYZ
                        : new double[] { 0, 0, 0 };

                    sb.AppendLine("    <inertial>");
                    sb.AppendLine("      <origin xyz=\"" + FormatXYZ(com) + "\" rpy=\"0 0 0\"/>");
                    sb.AppendLine("      <mass value=\"" + F(mass) + "\"/>");

                    double ixx = (Math.Abs(link.InertiaIxx) > 1e-20) ? link.InertiaIxx : 1e-6;
                    double iyy = (Math.Abs(link.InertiaIyy) > 1e-20) ? link.InertiaIyy : 1e-6;
                    double izz = (Math.Abs(link.InertiaIzz) > 1e-20) ? link.InertiaIzz : 1e-6;

                    sb.Append("      <inertia ");
                    sb.Append("ixx=\"").Append(F(ixx)).Append("\" ");
                    sb.Append("ixy=\"").Append(F(link.InertiaIxy)).Append("\" ");
                    sb.Append("ixz=\"").Append(F(link.InertiaIxz)).Append("\" ");
                    sb.Append("iyy=\"").Append(F(iyy)).Append("\" ");
                    sb.Append("iyz=\"").Append(F(link.InertiaIyz)).Append("\" ");
                    sb.Append("izz=\"").Append(F(izz)).Append("\"/>\n");
                    sb.AppendLine("    </inertial>");

                    // visual/collision
                    if (!string.IsNullOrEmpty(link.MeshFile))
                    {
                        sb.AppendLine("    <visual>");
                        sb.AppendLine("      <origin xyz=\"0 0 0\" rpy=\"0 0 0\"/>");
                        sb.AppendLine("      <geometry>");
                        sb.AppendLine("        <mesh filename=\"" + XmlEscape(link.MeshFile) + "\"/>");
                        sb.AppendLine("      </geometry>");
                        sb.AppendLine("      <material name=\"mat_" + XmlEscape(link.Name) + "\">");
                        sb.AppendLine("        <texture filename=\"meshes/" + XmlEscape(link.Name) + ".png\"/>");
                        sb.AppendLine("      </material>");
                        sb.AppendLine("    </visual>");

                        sb.AppendLine("    <collision>");
                        sb.AppendLine("      <origin xyz=\"0 0 0\" rpy=\"0 0 0\"/>");
                        sb.AppendLine("      <geometry>");
                        sb.AppendLine("        <mesh filename=\"" + XmlEscape(link.MeshFile) + "\"/>");
                        sb.AppendLine("      </geometry>");
                        sb.AppendLine("    </collision>");
                    }

                    sb.AppendLine("  </link>");
                }

                // ----------------
                // JOINTS
                // ----------------
                foreach (UrdfJoint j in validJoints)
                {
                    string type = string.IsNullOrEmpty(j.Type) ? "fixed" : j.Type;

                    sb.AppendLine("  <joint name=\"" + XmlEscape(j.Name) + "\" type=\"" + XmlEscape(type) + "\">");
                    sb.AppendLine("    <parent link=\"" + XmlEscape(j.ParentLink) + "\"/>");
                    sb.AppendLine("    <child link=\"" + XmlEscape(j.ChildLink) + "\"/>");

                    double[] oxyz = (j.OriginXYZ != null && j.OriginXYZ.Length >= 3) ? j.OriginXYZ : new double[] { 0, 0, 0 };
                    double[] orpy = (j.OriginRPY != null && j.OriginRPY.Length >= 3) ? j.OriginRPY : new double[] { 0, 0, 0 };
                    sb.AppendLine("    <origin xyz=\"" + FormatXYZ(oxyz) + "\" rpy=\"" + FormatXYZ(orpy) + "\"/>");

                    if (!string.Equals(type, "fixed", StringComparison.OrdinalIgnoreCase))
                    {
                        double[] axis = (j.AxisXYZ != null && j.AxisXYZ.Length == 3) ? j.AxisXYZ : new double[] { 0, 0, 1 };
                        axis = Normalize3(axis) ?? new double[] { 0, 0, 1 };
                        sb.AppendLine("    <axis xyz=\"" + FormatXYZ(axis) + "\"/>");

                        if (string.Equals(type, "revolute", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(type, "prismatic", StringComparison.OrdinalIgnoreCase))
                        {
                            double lower = j.LimitLower;
                            double upper = j.LimitUpper;

                            double effort = (j.LimitEffort > 0) ? j.LimitEffort : 10.0;
                            double vel = (j.LimitVelocity > 0) ? j.LimitVelocity : 1.0;

                            // si no hay límites, poner algo usable
                            if (Math.Abs(upper - lower) < 1e-12)
                            {
                                if (string.Equals(type, "revolute", StringComparison.OrdinalIgnoreCase))
                                {
                                    lower = -Math.PI;
                                    upper = Math.PI;
                                }
                                else
                                {
                                    lower = -0.1;
                                    upper = 0.1;
                                }
                            }

                            sb.Append("    <limit lower=\"").Append(F(lower))
                              .Append("\" upper=\"").Append(F(upper))
                              .Append("\" effort=\"").Append(F(effort))
                              .Append("\" velocity=\"").Append(F(vel))
                              .Append("\"/>\n");
                        }
                    }

                    sb.AppendLine("  </joint>");
                }

                sb.AppendLine("</robot>");

                IOFile.WriteAllText(urdfPath, sb.ToString(), new UTF8Encoding(false));
                DebugLog("SYS", "WriteUrdfFile OK: " + urdfPath);
            }
            catch (Exception ex)
            {
                DebugLog("ERR", "WriteUrdfFile ERROR: " + ex.Message);
            }
        }

        private static string FormatXYZ(double[] v)
        {
            if (v == null || v.Length < 3) return "0 0 0";
            return F(v[0]) + " " + F(v[1]) + " " + F(v[2]);
        }

        private static void EnsureBaseLink(RobotModel robot)
        {
            if (robot == null) return;

            UrdfLink baseLink = FindLinkByName(robot, "base_link");
            if (baseLink != null) return;

            baseLink = new UrdfLink();
            baseLink.Name = "base_link";
            baseLink.OriginXYZ = new double[] { 0, 0, 0 };
            baseLink.OriginRPY = new double[] { 0, 0, 0 };
            robot.Links.Insert(0, baseLink);

            DebugLog("LINK", "EnsureBaseLink: creado 'base_link'.");
        }

        private static void EnsureEveryLinkHasParentJoint(RobotModel robot)
        {
            if (robot == null || robot.Links == null) return;

            HashSet<string> hasParent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (UrdfJoint j in robot.Joints)
            {
                if (j == null) continue;
                if (!string.IsNullOrEmpty(j.ChildLink))
                    hasParent.Add(j.ChildLink);
            }

            foreach (UrdfLink link in robot.Links)
            {
                if (link == null || string.IsNullOrEmpty(link.Name)) continue;
                if (string.Equals(link.Name, "base_link", StringComparison.OrdinalIgnoreCase)) continue;

                if (hasParent.Contains(link.Name)) continue;

                // crear joint fijo a base_link con el origin del link (si existe)
                UrdfJoint jfix = new UrdfJoint();
                jfix.Name = "root_" + link.Name;
                jfix.Type = "fixed";
                jfix.ParentLink = "base_link";
                jfix.ChildLink = link.Name;

                double[] xyz = (link.OriginXYZ != null && link.OriginXYZ.Length >= 3) ? link.OriginXYZ : new double[] { 0, 0, 0 };
                double[] rpy = (link.OriginRPY != null && link.OriginRPY.Length >= 3) ? link.OriginRPY : new double[] { 0, 0, 0 };
                jfix.OriginXYZ = new double[] { xyz[0], xyz[1], xyz[2] };
                jfix.OriginRPY = new double[] { rpy[0], rpy[1], rpy[2] };

                robot.Joints.Add(jfix);
                hasParent.Add(link.Name);

                DebugLog("LINK", "EnsureEveryLinkHasParentJoint: añadido joint fijo '" + jfix.Name + "' para '" + link.Name + "'.");
            }
        }

        // =====================================================
        //  Matrix -> RPY (rad)  (roll X, pitch Y, yaw Z)
        // =====================================================
        private static void MatrixToRPY(Matrix m, out double roll, out double pitch, out double yaw)
        {
            roll = pitch = yaw = 0.0;
            if (m == null) return;

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
        }
    } // <-- fin UrdfExporter

    // ========================================================================
    //  MODELOS SIMPLES
    // ========================================================================
    public class RobotModel
    {
        public string Name = "robot";
        public List<UrdfLink> Links = new List<UrdfLink>();
        public List<UrdfJoint> Joints = new List<UrdfJoint>();
    }

    public class UrdfLink
    {
        public string Name;

        // Nota: OriginXYZ/RPY se usan como “hint” para joints root si faltan.
        public double[] OriginXYZ = new double[] { 0, 0, 0 };
        public double[] OriginRPY = new double[] { 0, 0, 0 };

        public string MeshFile;

        // Inertial (URDF)
        public double InertialMass = 0.01;
        public double[] InertialOriginXYZ = new double[] { 0, 0, 0 };
        public double[] InertialOriginRPY = new double[] { 0, 0, 0 };

        public double InertiaIxx = 1e-6;
        public double InertiaIxy = 0.0;
        public double InertiaIxz = 0.0;
        public double InertiaIyy = 1e-6;
        public double InertiaIyz = 0.0;
        public double InertiaIzz = 1e-6;

        // Aliases (por si en algún punto usaste reflection helpers)
        public double Mass { get { return InertialMass; } set { InertialMass = value; } }
        public double[] COM { get { return InertialOriginXYZ; } set { InertialOriginXYZ = value; } }
        public double Ixx { get { return InertiaIxx; } set { InertiaIxx = value; } }
        public double Iyy { get { return InertiaIyy; } set { InertiaIyy = value; } }
        public double Izz { get { return InertiaIzz; } set { InertiaIzz = value; } }
        public double Ixy { get { return InertiaIxy; } set { InertiaIxy = value; } }
        public double Ixz { get { return InertiaIxz; } set { InertiaIxz = value; } }
        public double Iyz { get { return InertiaIyz; } set { InertiaIyz = value; } }
    }

    public class UrdfJoint
    {
        public string Name;
        public string Type; // fixed / revolute / continuous / prismatic

        public string ParentLink;
        public string ChildLink;

        public double[] OriginXYZ = new double[] { 0, 0, 0 };
        public double[] OriginRPY = new double[] { 0, 0, 0 };

        public double[] AxisXYZ = null;

        public double LimitLower = 0.0;
        public double LimitUpper = 0.0;
        public double LimitEffort = 0.0;
        public double LimitVelocity = 0.0;
    }
} // <-- fin namespace
