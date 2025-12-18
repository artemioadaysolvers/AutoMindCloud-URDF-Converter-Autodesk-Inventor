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
                        AddInClientId,                       // ClientId = GUID del AddIn con llaves
                        "Export URDF with VeryLowOptimized mesh",
                        "Export URDF (VLQ)");
                }

                // Evitar múltiples suscripciones si Inventor re-llama Activate en la misma sesión
                try { _exportUrdfVlqButton.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfVlqButtonPressed); } catch { }
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
                        AddInClientId,                       // MISMO ClientId
                        "Export URDF with DisplayMesh-quality mesh",
                        "Export URDF (Display)");
                }

                try { _exportUrdfDisplayButton.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed); } catch { }
                _exportUrdfDisplayButton.OnExecute +=
                    new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed);

                Debug.WriteLine("[URDF][SYS] ButtonDefinitions creados/obtenidos correctamente.");

                // -------------------------------------------------
                // 2) Añadir los botones a los Ribbons de Part y Assembly
                // -------------------------------------------------
                UserInterfaceManager uiMgr = _invApp.UserInterfaceManager;

                Debug.WriteLine("[URDF][SYS] Creando panels en ribbons...");

                // 2.1) Ribbon de PIEZAS (Part)
                try
                {
                    Ribbon partRibbon = uiMgr.Ribbons["Part"];
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
                            AddInClientId,   // ✅ antes tenías "aaaaaaaa..." (mal). Debe ser tu ClientId real.
                            "",
                            false);
                    }

                    // Añadimos los botones al panel (sin duplicar)
                    SafeAddButtonToPanel(urdfPanelPart, _exportUrdfVlqButton);
                    SafeAddButtonToPanel(urdfPanelPart, _exportUrdfDisplayButton);

                    Debug.WriteLine("[URDF][UI] Panel URDF en Part creado/actualizado correctamente.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[URDF][UI] Error creando panel en Part: " + ex.Message);
                }

                // 2.2) Ribbon de ENSAMBLAJES (Assembly)
                try
                {
                    Ribbon asmRibbon = uiMgr.Ribbons["Assembly"];
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
                            AddInClientId,   // ✅ antes tenías "bbbb..." (mal). Debe ser tu ClientId real.
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
                // Evitar duplicados si Activate corre más de una vez
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

            try
            {
                panel.CommandControls.AddButton(btn);
            }
            catch { }
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

            string baseDir = IOPath.GetDirectoryName(fullPath);
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
        // =====================================================
        private static void AddAssemblyOccurrencesAndBodiesAsLinks(
            AssemblyDocument asmDoc,
            RobotModel robot)
        {
            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs = asmDef.Occurrences;

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

                    string rawName = occ.Name;
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

                            DebugLog("LINK",
                                "Añadido link principal '" + linkName +
                                "' colgando de base_link.");
                        }
                        else
                        {
                            joint.Name = "fixed_extra_" + linkName;
                            joint.ParentLink = baseLinkName;
                            joint.ChildLink = linkName;
                            joint.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                            joint.OriginRPY = new double[] { 0.0, 0.0, 0.0 };

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
                DebugLog("MESH",
                        "TryGetColorFromNamedAssetValue: error buscando '" + targetName + "'.");
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

            // Log completo
            LogAssetInfo(ownerKind, ownerName, app);

            // 1) plasticvinyl_color
            if (TryGetColorFromNamedAssetValue(app, "plasticvinyl_color", out r, out g, out b))
            {
                DebugLog("MESH",
                        "TryGetColorFromAssetWithPriority: usando plasticvinyl_color para " +
                        ownerKind + "='" + ownerName + "'.");
                return true;
            }

            // 2) wallpaint_color
            if (TryGetColorFromNamedAssetValue(app, "wallpaint_color", out r, out g, out b))
            {
                DebugLog("MESH",
                        "TryGetColorFromAssetWithPriority: usando wallpaint_color para " +
                        ownerKind + "='" + ownerName + "'.");
                return true;
            }

            // 3) generic_diffuse (ANTES de common_Tint_color para que salga el rojo en PNG)
            try
            {
                // Primero intenta por helper (si lo soporta)
                if (TryGetColorFromNamedAssetValue(app, "generic_diffuse", out r, out g, out b))
                {
                    DebugLog("MESH",
                            "TryGetColorFromAssetWithPriority: usando generic_diffuse para " +
                            ownerKind + "='" + ownerName + "'.");
                    return true;
                }

                // Fallback manual (por si el helper no entiende generic_diffuse)
                AssetValue avDif = null;
                try { avDif = app["generic_diffuse"]; } catch { avDif = null; }

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

            // 4) common_Tint_color (solo si no hubo diffuse)
            //    Además, si es gris (típico 80,80,80) no retornamos, para no pisar colores reales.
            try
            {
                double tr, tg, tb;
                if (TryGetColorFromNamedAssetValue(app, "common_Tint_color", out tr, out tg, out tb))
                {
                    bool isGrayish =
                            Math.Abs(tr - tg) < 0.02 &&
                            Math.Abs(tg - tb) < 0.02;

                    if (!isGrayish)
                    {
                        r = tr;
                        g = tg;
                        b = tb;

                        DebugLog("MESH",
                                "TryGetColorFromAssetWithPriority: usando common_Tint_color para " +
                                ownerKind + "='" + ownerName + "'.");
                        return true;
                    }

                    DebugLog("MESH",
                            "TryGetColorFromAssetWithPriority: common_Tint_color es grisáceo, se ignora para " +
                            ownerKind + "='" + ownerName + "'.");
                }
            }
            catch
            {
                // no-op, seguimos con el fallback general
            }

            // 5) primer AssetValue COLOR
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
                                r = invCol.Red / 255.0;
                                g = invCol.Green / 255.0;
                                b = invCol.Blue / 255.0;

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
        // =====================================================

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
                try { appBody = body.Appearance; }
                catch { appBody = null; }

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

            // 2) Fallback: Occurrence.Appearance
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

        private static bool TryGetFaceColor(
                Inventor.Face face,
                SurfaceBody parentBody,
                string ownerNameForLog,
                Asset occAppearance,
                out double r,
                out double g,
                out double b)
        {
            r = 0.8;
            g = 0.8;
            b = 0.8;

            if (string.IsNullOrEmpty(ownerNameForLog))
                ownerNameForLog = "(Face)";

            // 1) Face.Appearance
            if (face != null)
            {
                try
                {
                    Asset app = null;
                    try { app = face.Appearance; }
                    catch { app = null; }

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

            // 2) Fallback: Body/Occurrence
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
            int width = cellsX * cellSize;
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
            try { faces = body.Faces; }
            catch { faces = null; }

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

            int width = cellsX * cellSize;
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
                    try { f = faces[fi + 1]; }
                    catch { f = null; }

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

            DebugLog("MESH", "WriteBodyFaceColorAtlasCore: atlas escrito OK en '" + path + "'");
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

                string daeName = linkName + ".dae";
                string daePath = IOPath.Combine(meshesDir, daeName);

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
                    WriteSolidColorPng(pngPath, r, g, b, 32);
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: LOW PNG='" + pngPath + "'");
                }
                else if (_meshQualityMode == "high")
                {
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

            // 🔧 Ajustar tipos de JOINT + AXIS + PARENT + ORIGIN (robusto)
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

                    string rawName = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    string baseLinkName = "link_" +
                                          occIndex.ToString(CultureInfo.InvariantCulture) +
                                          "_" + safeName;

                    // Apariencia a nivel de OCCURRENCE
                    Asset occAppearance = null;
                    try { occAppearance = occ.Appearance; }
                    catch { occAppearance = null; }

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
                        double r, g, b;
                        if (!TryGetBodyColor(body, rawName, occAppearance, out r, out g, out b))
                        {
                            r = g = b = 0.8;
                        }

                        string pngPath = IOPath.Combine(meshesDir, linkName + ".png");

                        if (_meshQualityMode == "low")
                        {
                            WriteSolidColorPng(pngPath, r, g, b, 32);
                            DebugLog("MESH",
                                "ExportAssemblyGeometryToDae: LOW PNG='" + pngPath + "'");
                        }
                        else if (_meshQualityMode == "high")
                        {
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
        //  Resolver occurrences de constraints a una leaf occurrence mapeada
        //  (si el constraint apunta a un sub-assembly occurrence)
        // -------------------------------------------------
        private static ComponentOccurrence ResolveToMappedLeafOccurrence(
            ComponentOccurrence occ,
            Dictionary<ComponentOccurrence, string> occToBaseLink)
        {
            if (occ == null) return null;

            if (occToBaseLink != null && occToBaseLink.ContainsKey(occ))
                return occ;

            try
            {
                // Inventor: SubOccurrences es ComponentOccurrencesEnumerator (NO ComponentOccurrences)
                ComponentOccurrencesEnumerator subs = null;
                try { subs = occ.SubOccurrences; } catch { subs = null; }

                if (subs != null)
                {
                    foreach (ComponentOccurrence so in subs)
                    {
                        if (so == null) continue;

                        if (occToBaseLink != null && occToBaseLink.ContainsKey(so))
                            return so;

                        // Profundizar (por si so no es leaf pero tiene descendientes leaf mapeados)
                        ComponentOccurrence deep = ResolveToMappedLeafOccurrence(so, occToBaseLink);
                        if (deep != null && occToBaseLink != null && occToBaseLink.ContainsKey(deep))
                            return deep;
                    }
                }
            }
            catch { }

            // No encontramos un leaf mapeado: devolvemos el mismo occ como fallback
            return occ;
        }


        // -------------------------------------------------
        //  Mapear AssemblyConstraints → tipos de JOINT URDF (ROBUSTO)
        //
        //  ✅ No asume que child = OccurrenceOne.
        //  ✅ Detecta MateConstraint entre cilindros / workaxis / edge lineal → continuous.
        //  ✅ Extrae eje REAL desde GeometryOne/GeometryTwo (en espacio del assembly).
        //  ✅ Convierte axis(world) → axis(joint frame) usando la rotación de la occurrence del child.
        //  ✅ AHORA TAMBIÉN: setea ParentLink real y Origin RELATIVO (parent→child) cuando se puede.
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

            // Mapeo robusto: ComponentOccurrence → baseLinkName
            Dictionary<ComponentOccurrence, string> occToBaseLink =
                new Dictionary<ComponentOccurrence, string>();

            // Guardamos Matrix por occurrence (para axis)
            Dictionary<ComponentOccurrence, Matrix> occToMatrix =
                new Dictionary<ComponentOccurrence, Matrix>();

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

                    string rawName = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    string baseLinkName = "link_" +
                        occIndex.ToString(CultureInfo.InvariantCulture) +
                        "_" + safeName;

                    occToBaseLink[occ] = baseLinkName;

                    try { occToMatrix[occ] = occ.Transformation; }
                    catch { }

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
            try { constraints = asmDef.Constraints; }
            catch { constraints = null; }

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

                ComponentOccurrence o1 = null;
                ComponentOccurrence o2 = null;

                try
                {
                    o1 = ac.OccurrenceOne;
                    o2 = ac.OccurrenceTwo;
                }
                catch
                {
                    o1 = null;
                    o2 = null;
                }

                // Resolver a leaf mapeada si venía un sub-assembly occurrence
                ComponentOccurrence ro1 = ResolveToMappedLeafOccurrence(o1, occToBaseLink);
                ComponentOccurrence ro2 = ResolveToMappedLeafOccurrence(o2, occToBaseLink);

                string link1 = null, link2 = null;
                if (ro1 != null) occToBaseLink.TryGetValue(ro1, out link1);
                if (ro2 != null) occToBaseLink.TryGetValue(ro2, out link2);

                // Si no tenemos ninguno (no-leaf / suprimido), no podemos mapear
                if (string.IsNullOrEmpty(link1) && string.IsNullOrEmpty(link2))
                    continue;

                // Elegir child robusto (NO asumir OccurrenceOne)
                ComponentOccurrence childOcc = null;
                ComponentOccurrence parentOcc = null;
                string childLink = null;
                string parentLink = null;

                bool g1 = false, g2 = false;
                try { if (ro1 != null) g1 = ro1.Grounded; } catch { g1 = false; }
                try { if (ro2 != null) g2 = ro2.Grounded; } catch { g2 = false; }

                if (!string.IsNullOrEmpty(link1) && !string.IsNullOrEmpty(link2))
                {
                    if (g1 && !g2)
                    {
                        parentOcc = ro1;
                        childOcc = ro2;
                        parentLink = link1;
                        childLink = link2;
                    }
                    else if (g2 && !g1)
                    {
                        parentOcc = ro2;
                        childOcc = ro1;
                        parentLink = link2;
                        childLink = link1;
                    }
                    else
                    {
                        // Si ambos grounded o ambos libres:
                        // preferimos modificar el que todavía tenga joint FIXED (para no pisar un DOF ya asignado)
                        UrdfJoint j1 = FindJointByChildLink(robot, link1);
                        UrdfJoint j2 = FindJointByChildLink(robot, link2);

                        bool j1Fixed = (j1 != null && string.Equals(j1.Type, "fixed", StringComparison.OrdinalIgnoreCase));
                        bool j2Fixed = (j2 != null && string.Equals(j2.Type, "fixed", StringComparison.OrdinalIgnoreCase));

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
                    // Solo uno está en el mapa leaf:
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
                if (joint == null)
                {
                    DebugLog("LINK",
                        "UpdateJointTypesFromConstraints: no se encontró JOINT para childLink='" +
                        childLink + "' (constraint '" + ac.Name + "').");
                    continue;
                }

                // Solo reasignar si estaba FIXED (no pisar)
                if (!string.Equals(joint.Type, "fixed", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string newType = null;

                // Intentar extraer eje REAL (world) desde GeometryOne/Two o EntityOne/Two
                double[] axisWorld = null;
                string axisSrc = null;
                bool gotAxis = TryExtractAxisWorldFromConstraint(ac, out axisWorld, out axisSrc);

                // -------- Mapeo por tipo de constraint --------
                if (ac is InsertConstraint)
                {
                    newType = "continuous";
                }
                else if (ac is AngleConstraint)
                {
                    newType = "revolute";
                }
                else if (ac is TransitionalConstraint)
                {
                    newType = "prismatic";
                }
                else if (ac is MateConstraint)
                {
                    // SOLO Mate “eje” (cilindro / workaxis / edge lineal) → DOF rotacional
                    if (gotAxis) newType = "continuous";
                    else newType = null;
                }
                else
                {
                    newType = null;
                }

                if (string.IsNullOrEmpty(newType))
                    continue;

                joint.Type = newType;

                // ✅ Parent real si lo tenemos (si no, se queda como estaba)
                if (!string.IsNullOrEmpty(parentLink))
                {
                    // Evitar autoloops
                    if (!string.Equals(parentLink, joint.ChildLink, StringComparison.OrdinalIgnoreCase))
                        joint.ParentLink = parentLink;
                }

                // ✅ Origin relativo parent→child (si tenemos matrices de ambos)
                //    Esto es CLAVE para que al cambiar a continuous/revolute no “quede todo colgando de base_link”.
                try
                {
                    Matrix parentM = null;
                    Matrix childM = null;

                    if (parentOcc != null) occToMatrix.TryGetValue(parentOcc, out parentM);
                    if (childOcc != null) occToMatrix.TryGetValue(childOcc, out childM);

                    double tx_m, ty_m, tz_m, r, p, y;
                    if (TryComputeRelativeXYZRPY(parentM, childM, out tx_m, out ty_m, out tz_m, out r, out p, out y))
                    {
                        joint.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                        joint.OriginRPY = new double[] { r, p, y };

                        DebugLog("LINK",
                            "Constraint '" + ac.Name + "' → joint '" + joint.Name +
                            "' origin REL parent→child xyz(m)=(" +
                            tx_m.ToString("G6", CultureInfo.InvariantCulture) + "," +
                            ty_m.ToString("G6", CultureInfo.InvariantCulture) + "," +
                            tz_m.ToString("G6", CultureInfo.InvariantCulture) + ") rpy(rad)=(" +
                            r.ToString("G6", CultureInfo.InvariantCulture) + "," +
                            p.ToString("G6", CultureInfo.InvariantCulture) + "," +
                            y.ToString("G6", CultureInfo.InvariantCulture) + ")");
                    }
                }
                catch { }

                // Si el joint es DOF y tenemos axis, convertirlo al frame del joint (child occurrence)
                bool isFixed = string.Equals(joint.Type, "fixed", StringComparison.OrdinalIgnoreCase);
                if (!isFixed && gotAxis && axisWorld != null && axisWorld.Length == 3)
                {
                    Matrix childM = null;
                    occToMatrix.TryGetValue(childOcc, out childM);

                    double[] axisLocal = AxisWorldToOccLocal(axisWorld, childM);
                    if (axisLocal != null && axisLocal.Length == 3)
                    {
                        joint.AxisXYZ = axisLocal;

                        DebugLog("LINK",
                            "Constraint '" + ac.Name + "' → joint '" + joint.Name +
                            "' type='" + joint.Type +
                            "' axisSrc='" + (axisSrc ?? "unknown") +
                            "' axis(local)=(" +
                            axisLocal[0].ToString("G6", CultureInfo.InvariantCulture) + "," +
                            axisLocal[1].ToString("G6", CultureInfo.InvariantCulture) + "," +
                            axisLocal[2].ToString("G6", CultureInfo.InvariantCulture) + ")" +
                            " childOcc='" + (childOcc != null ? childOcc.Name : "(null)") + "'");
                    }
                }
                else
                {
                    DebugLog("LINK",
                        "Constraint '" + ac.Name + "' → joint '" + joint.Name +
                        "' type='" + joint.Type +
                        "' (sin axis real, queda default o no aplica).");
                }
            }
        }

        // =====================================================
        //  RELATIVE origin: parent→child (m) + RPY(rad)
        //  Matrices de Inventor están en cm y son rígidas (rot+trasl).
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

            if (parentM == null || childM == null)
                return false;

            // Parent R (world)
            double pr11 = parentM.Cell[1, 1];
            double pr12 = parentM.Cell[1, 2];
            double pr13 = parentM.Cell[1, 3];
            double pr21 = parentM.Cell[2, 1];
            double pr22 = parentM.Cell[2, 2];
            double pr23 = parentM.Cell[2, 3];
            double pr31 = parentM.Cell[3, 1];
            double pr32 = parentM.Cell[3, 2];
            double pr33 = parentM.Cell[3, 3];

            // Child R (world)
            double cr11 = childM.Cell[1, 1];
            double cr12 = childM.Cell[1, 2];
            double cr13 = childM.Cell[1, 3];
            double cr21 = childM.Cell[2, 1];
            double cr22 = childM.Cell[2, 2];
            double cr23 = childM.Cell[2, 3];
            double cr31 = childM.Cell[3, 1];
            double cr32 = childM.Cell[3, 2];
            double cr33 = childM.Cell[3, 3];

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

            // cm → m
            tx_m = tx_cm * 0.01;
            ty_m = ty_cm * 0.01;
            tz_m = tz_cm * 0.01;

            // R → RPY (mismo esquema que MatrixToRPY)
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

            if (TryGetAxisFromAnyObject(g1, out axisWorld, out axisSource))
                return true;
            if (TryGetAxisFromAnyObject(g2, out axisWorld, out axisSource))
                return true;

            if (TryGetAxisFromAnyObject(e1, out axisWorld, out axisSource))
                return true;
            if (TryGetAxisFromAnyObject(e2, out axisWorld, out axisSource))
                return true;

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
        //  axis_local = R^T * axis_world   (sin traslación)
        // =====================================================
        private static double[] AxisWorldToOccLocal(double[] axisWorld, Matrix occMatrix)
        {
            if (axisWorld == null || axisWorld.Length != 3)
                return null;

            double[] a = Normalize3(new double[] { axisWorld[0], axisWorld[1], axisWorld[2] });
            if (a == null) return null;

            if (occMatrix == null)
                return a;

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
        //  Continúa en BLOQUE 2/2: Collada, Inertial, URDF writer,
        //  MatrixToRPY, clases RobotModel/UrdfLink/UrdfJoint, etc.
        // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>










        // =====================================================
        //  COLLADA (.dae) writer: geometría + UV + textura (PNG)
        // =====================================================
        private static void WriteColladaFile(
            string daePath,
            string meshId,
            double[] vertices,   // xyz en metros
            int[] indices)       // tri list, 0-based
        {
            if (string.IsNullOrEmpty(daePath) || vertices == null || indices == null)
                return;

            int vCount = vertices.Length / 3;
            int triCount = indices.Length / 3;

            // Normales por vértice (aprox) y UV planar (bbox XY)
            double[] normals = BuildVertexNormals(vertices, indices);
            double[] uvs = BuildPlanarUV(vertices);

            // Textura: asumimos <meshId>.png en la misma carpeta que el .dae
            string pngFileName = meshId + ".png";

            string geomId = meshId + "-geom";
            string posId = geomId + "-positions";
            string norId = geomId + "-normals";
            string uvId  = geomId + "-uvs";
            string vtxId = geomId + "-vertices";

            string imgId = meshId + "-img";
            string effId = meshId + "-effect";
            string matId = meshId + "-material";

            StringBuilder sb = new StringBuilder(1024 * 64);

            sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.AppendLine("<COLLADA xmlns=\"http://www.collada.org/2005/11/COLLADASchema\" version=\"1.4.1\">");
            sb.AppendLine("  <asset>");
            sb.AppendLine("    <contributor><authoring_tool>URDFConverterAddIn</authoring_tool></contributor>");
            sb.AppendLine("    <unit name=\"meter\" meter=\"1\"/>");
            sb.AppendLine("    <up_axis>Z_UP</up_axis>");
            sb.AppendLine("  </asset>");

            // Images
            sb.AppendLine("  <library_images>");
            sb.AppendLine("    <image id=\"" + XmlEscape(imgId) + "\" name=\"" + XmlEscape(imgId) + "\">");
            sb.AppendLine("      <init_from>" + XmlEscape(pngFileName) + "</init_from>");
            sb.AppendLine("    </image>");
            sb.AppendLine("  </library_images>");

            // Effects (lambert + texture)
            sb.AppendLine("  <library_effects>");
            sb.AppendLine("    <effect id=\"" + XmlEscape(effId) + "\">");
            sb.AppendLine("      <profile_COMMON>");
            sb.AppendLine("        <newparam sid=\"surface\">");
            sb.AppendLine("          <surface type=\"2D\">");
            sb.AppendLine("            <init_from>" + XmlEscape(imgId) + "</init_from>");
            sb.AppendLine("          </surface>");
            sb.AppendLine("        </newparam>");
            sb.AppendLine("        <newparam sid=\"sampler\">");
            sb.AppendLine("          <sampler2D><source>surface</source></sampler2D>");
            sb.AppendLine("        </newparam>");
            sb.AppendLine("        <technique sid=\"common\">");
            sb.AppendLine("          <lambert>");
            sb.AppendLine("            <diffuse>");
            sb.AppendLine("              <texture texture=\"sampler\" texcoord=\"UVSET0\"/>");
            sb.AppendLine("            </diffuse>");
            sb.AppendLine("          </lambert>");
            sb.AppendLine("        </technique>");
            sb.AppendLine("      </profile_COMMON>");
            sb.AppendLine("    </effect>");
            sb.AppendLine("  </library_effects>");

            // Materials
            sb.AppendLine("  <library_materials>");
            sb.AppendLine("    <material id=\"" + XmlEscape(matId) + "\" name=\"" + XmlEscape(matId) + "\">");
            sb.AppendLine("      <instance_effect url=\"#" + XmlEscape(effId) + "\"/>");
            sb.AppendLine("    </material>");
            sb.AppendLine("  </library_materials>");

            // Geometries
            sb.AppendLine("  <library_geometries>");
            sb.AppendLine("    <geometry id=\"" + XmlEscape(geomId) + "\" name=\"" + XmlEscape(meshId) + "\">");
            sb.AppendLine("      <mesh>");

            // Positions
            sb.AppendLine("        <source id=\"" + XmlEscape(posId) + "\">");
            sb.AppendLine("          <float_array id=\"" + XmlEscape(posId) + "-array\" count=\"" + (vCount * 3).ToString(CultureInfo.InvariantCulture) + "\">");
            for (int i = 0; i < vCount; i++)
            {
                int k = i * 3;
                sb.Append(FormatF(vertices[k])).Append(" ")
                  .Append(FormatF(vertices[k + 1])).Append(" ")
                  .Append(FormatF(vertices[k + 2])).Append(i == vCount - 1 ? "" : " ");
            }
            sb.AppendLine();
            sb.AppendLine("          </float_array>");
            sb.AppendLine("          <technique_common>");
            sb.AppendLine("            <accessor source=\"#" + XmlEscape(posId) + "-array\" count=\"" + vCount.ToString(CultureInfo.InvariantCulture) + "\" stride=\"3\">");
            sb.AppendLine("              <param name=\"X\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"Y\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"Z\" type=\"float\"/>");
            sb.AppendLine("            </accessor>");
            sb.AppendLine("          </technique_common>");
            sb.AppendLine("        </source>");

            // Normals
            sb.AppendLine("        <source id=\"" + XmlEscape(norId) + "\">");
            sb.AppendLine("          <float_array id=\"" + XmlEscape(norId) + "-array\" count=\"" + (vCount * 3).ToString(CultureInfo.InvariantCulture) + "\">");
            for (int i = 0; i < vCount; i++)
            {
                int k = i * 3;
                sb.Append(FormatF(normals[k])).Append(" ")
                  .Append(FormatF(normals[k + 1])).Append(" ")
                  .Append(FormatF(normals[k + 2])).Append(i == vCount - 1 ? "" : " ");
            }
            sb.AppendLine();
            sb.AppendLine("          </float_array>");
            sb.AppendLine("          <technique_common>");
            sb.AppendLine("            <accessor source=\"#" + XmlEscape(norId) + "-array\" count=\"" + vCount.ToString(CultureInfo.InvariantCulture) + "\" stride=\"3\">");
            sb.AppendLine("              <param name=\"X\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"Y\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"Z\" type=\"float\"/>");
            sb.AppendLine("            </accessor>");
            sb.AppendLine("          </technique_common>");
            sb.AppendLine("        </source>");

            // UVs
            sb.AppendLine("        <source id=\"" + XmlEscape(uvId) + "\">");
            sb.AppendLine("          <float_array id=\"" + XmlEscape(uvId) + "-array\" count=\"" + (vCount * 2).ToString(CultureInfo.InvariantCulture) + "\">");
            for (int i = 0; i < vCount; i++)
            {
                int k = i * 2;
                sb.Append(FormatF(uvs[k])).Append(" ")
                  .Append(FormatF(uvs[k + 1])).Append(i == vCount - 1 ? "" : " ");
            }
            sb.AppendLine();
            sb.AppendLine("          </float_array>");
            sb.AppendLine("          <technique_common>");
            sb.AppendLine("            <accessor source=\"#" + XmlEscape(uvId) + "-array\" count=\"" + vCount.ToString(CultureInfo.InvariantCulture) + "\" stride=\"2\">");
            sb.AppendLine("              <param name=\"S\" type=\"float\"/>");
            sb.AppendLine("              <param name=\"T\" type=\"float\"/>");
            sb.AppendLine("            </accessor>");
            sb.AppendLine("          </technique_common>");
            sb.AppendLine("        </source>");

            // Vertices
            sb.AppendLine("        <vertices id=\"" + XmlEscape(vtxId) + "\">");
            sb.AppendLine("          <input semantic=\"POSITION\" source=\"#" + XmlEscape(posId) + "\"/>");
            sb.AppendLine("        </vertices>");

            // Triangles (offsets: 0=VERTEX, 1=NORMAL, 2=TEXCOORD)
            sb.AppendLine("        <triangles material=\"" + XmlEscape(matId) + "\" count=\"" + triCount.ToString(CultureInfo.InvariantCulture) + "\">");
            sb.AppendLine("          <input semantic=\"VERTEX\" source=\"#" + XmlEscape(vtxId) + "\" offset=\"0\"/>");
            sb.AppendLine("          <input semantic=\"NORMAL\" source=\"#" + XmlEscape(norId) + "\" offset=\"1\"/>");
            sb.AppendLine("          <input semantic=\"TEXCOORD\" source=\"#" + XmlEscape(uvId) + "\" offset=\"2\" set=\"0\"/>");
            sb.Append("          <p>");
            for (int t = 0; t < triCount; t++)
            {
                int i0 = indices[t * 3 + 0];
                int i1 = indices[t * 3 + 1];
                int i2 = indices[t * 3 + 2];

                // Para cada vértice: vIndex nIndex uvIndex (todos iguales, 1:1)
                sb.Append(i0.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(i0.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(i0.ToString(CultureInfo.InvariantCulture)).Append(" ");

                sb.Append(i1.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(i1.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(i1.ToString(CultureInfo.InvariantCulture)).Append(" ");

                sb.Append(i2.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(i2.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(i2.ToString(CultureInfo.InvariantCulture));

                if (t != triCount - 1) sb.Append(" ");
            }
            sb.AppendLine("</p>");
            sb.AppendLine("        </triangles>");

            sb.AppendLine("      </mesh>");
            sb.AppendLine("    </geometry>");
            sb.AppendLine("  </library_geometries>");

            // Visual Scene
            sb.AppendLine("  <library_visual_scenes>");
            sb.AppendLine("    <visual_scene id=\"Scene\" name=\"Scene\">");
            sb.AppendLine("      <node id=\"" + XmlEscape(meshId) + "\" name=\"" + XmlEscape(meshId) + "\">");
            sb.AppendLine("        <instance_geometry url=\"#" + XmlEscape(geomId) + "\">");
            sb.AppendLine("          <bind_material>");
            sb.AppendLine("            <technique_common>");
            sb.AppendLine("              <instance_material symbol=\"" + XmlEscape(matId) + "\" target=\"#" + XmlEscape(matId) + "\">");
            sb.AppendLine("                <bind_vertex_input semantic=\"UVSET0\" input_semantic=\"TEXCOORD\" input_set=\"0\"/>");
            sb.AppendLine("              </instance_material>");
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

            IOFile.WriteAllText(daePath, sb.ToString(), Encoding.UTF8);
        }

        private static double[] BuildVertexNormals(double[] vertices, int[] indices)
        {
            int vCount = vertices.Length / 3;
            double[] n = new double[vCount * 3];

            int triCount = indices.Length / 3;
            for (int t = 0; t < triCount; t++)
            {
                int i0 = indices[t * 3 + 0];
                int i1 = indices[t * 3 + 1];
                int i2 = indices[t * 3 + 2];

                int a0 = i0 * 3, a1 = i1 * 3, a2 = i2 * 3;

                double x0 = vertices[a0], y0 = vertices[a0 + 1], z0 = vertices[a0 + 2];
                double x1 = vertices[a1], y1 = vertices[a1 + 1], z1 = vertices[a1 + 2];
                double x2 = vertices[a2], y2 = vertices[a2 + 1], z2 = vertices[a2 + 2];

                double ux = x1 - x0, uy = y1 - y0, uz = z1 - z0;
                double vx = x2 - x0, vy = y2 - y0, vz = z2 - z0;

                double nx = uy * vz - uz * vy;
                double ny = uz * vx - ux * vz;
                double nz = ux * vy - uy * vx;

                // acumular
                n[a0] += nx; n[a0 + 1] += ny; n[a0 + 2] += nz;
                n[a1] += nx; n[a1 + 1] += ny; n[a1 + 2] += nz;
                n[a2] += nx; n[a2 + 1] += ny; n[a2 + 2] += nz;
            }

            for (int i = 0; i < vCount; i++)
            {
                int k = i * 3;
                double x = n[k], y = n[k + 1], z = n[k + 2];
                double len = Math.Sqrt(x * x + y * y + z * z);
                if (len < 1e-12)
                {
                    n[k] = 0; n[k + 1] = 0; n[k + 2] = 1;
                }
                else
                {
                    n[k] = x / len; n[k + 1] = y / len; n[k + 2] = z / len;
                }
            }

            return n;
        }

        private static double[] BuildPlanarUV(double[] vertices)
        {
            int vCount = vertices.Length / 3;
            double[] uv = new double[vCount * 2];

            double minX = double.PositiveInfinity, minY = double.PositiveInfinity;
            double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;

            for (int i = 0; i < vCount; i++)
            {
                int k = i * 3;
                double x = vertices[k];
                double y = vertices[k + 1];

                if (x < minX) minX = x;
                if (y < minY) minY = y;
                if (x > maxX) maxX = x;
                if (y > maxY) maxY = y;
            }

            double dx = maxX - minX;
            double dy = maxY - minY;
            if (dx < 1e-12) dx = 1.0;
            if (dy < 1e-12) dy = 1.0;

            for (int i = 0; i < vCount; i++)
            {
                int k3 = i * 3;
                int k2 = i * 2;

                double x = vertices[k3];
                double y = vertices[k3 + 1];

                double u = (x - minX) / dx;
                double v = (y - minY) / dy;

                // Collada/DAE normalmente usa (0,0) abajo-izquierda; dejamos tal cual
                uv[k2] = u;
                uv[k2 + 1] = v;
            }

            return uv;
        }

        private static string XmlEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;").Replace("'", "&apos;");
        }

        private static string FormatF(double v)
        {
            // compacto pero estable
            return v.ToString("G9", CultureInfo.InvariantCulture);
        }

        // =====================================================
        //  INERTIAL: leer MassProperties con fallback robusto
        // =====================================================
        private static void FillLinkInertialFromMassProperties(UrdfLink link, MassProperties mp)
        {
            if (link == null || mp == null)
                return;

            // Defaults "seguros"
            link.InertialMass = 1.0;
            link.InertialOriginXYZ = new double[] { 0, 0, 0 };
            link.InertiaIxx = 1e-4;
            link.InertiaIyy = 1e-4;
            link.InertiaIzz = 1e-4;
            link.InertiaIxy = 0;
            link.InertiaIxz = 0;
            link.InertiaIyz = 0;

            try
            {
                // Mass (kg)
                try
                {
                    link.InertialMass = mp.Mass;
                }
                catch { }

                // Center of mass (Inventor: cm)
                try
                {
                    object comObj = null;
                    try { comObj = mp.CenterOfMass; } catch { comObj = null; }

                    if (comObj != null)
                    {
                        // Inventor.Point normalmente tiene X/Y/Z en cm
                        double cx = 0, cy = 0, cz = 0;

                        try { cx = (double)comObj.GetType().InvokeMember("X", System.Reflection.BindingFlags.GetProperty, null, comObj, null); } catch { }
                        try { cy = (double)comObj.GetType().InvokeMember("Y", System.Reflection.BindingFlags.GetProperty, null, comObj, null); } catch { }
                        try { cz = (double)comObj.GetType().InvokeMember("Z", System.Reflection.BindingFlags.GetProperty, null, comObj, null); } catch { }

                        link.InertialOriginXYZ = new double[] { cx * 0.01, cy * 0.01, cz * 0.01 };
                    }
                }
                catch { }

                // Inertia tensor: intentamos varias rutas (depende versión/API)
                // Si viene en kg*cm^2, convertir a kg*m^2 multiplicando por 1e-4.
                bool gotTensor = false;

                try
                {
                    object tensor = null;
                    try { tensor = mp.GetType().InvokeMember("InertiaTensor", System.Reflection.BindingFlags.GetProperty, null, mp, null); }
                    catch { tensor = null; }

                    if (tensor != null)
                    {
                        // Algunos exponen Matrix con Cell[1..3,1..3]
                        try
                        {
                            // Inventor.Matrix tiene Cell[row,col]
                            double i11 = (double)tensor.GetType().InvokeMember("Cell", System.Reflection.BindingFlags.GetProperty, null, tensor, new object[] { 1, 1 });
                            double i22 = (double)tensor.GetType().InvokeMember("Cell", System.Reflection.BindingFlags.GetProperty, null, tensor, new object[] { 2, 2 });
                            double i33 = (double)tensor.GetType().InvokeMember("Cell", System.Reflection.BindingFlags.GetProperty, null, tensor, new object[] { 3, 3 });

                            double i12 = (double)tensor.GetType().InvokeMember("Cell", System.Reflection.BindingFlags.GetProperty, null, tensor, new object[] { 1, 2 });
                            double i13 = (double)tensor.GetType().InvokeMember("Cell", System.Reflection.BindingFlags.GetProperty, null, tensor, new object[] { 1, 3 });
                            double i23 = (double)tensor.GetType().InvokeMember("Cell", System.Reflection.BindingFlags.GetProperty, null, tensor, new object[] { 2, 3 });

                            double s = 1e-4; // cm^2 -> m^2
                            link.InertiaIxx = i11 * s;
                            link.InertiaIyy = i22 * s;
                            link.InertiaIzz = i33 * s;
                            link.InertiaIxy = i12 * s;
                            link.InertiaIxz = i13 * s;
                            link.InertiaIyz = i23 * s;

                            gotTensor = true;
                        }
                        catch { }
                    }
                }
                catch { }

                if (!gotTensor)
                {
                    // fallback: dejar diagonal pequeña proporcional a masa
                    double m = (link.InertialMass > 1e-9) ? link.InertialMass : 1.0;
                    double d = 1e-4 * m;
                    link.InertiaIxx = d;
                    link.InertiaIyy = d;
                    link.InertiaIzz = d;
                    link.InertiaIxy = 0;
                    link.InertiaIxz = 0;
                    link.InertiaIyz = 0;
                }
            }
            catch
            {
                // Si algo falla: se quedan defaults
            }
        }

        // =====================================================
        //  URDF writer
        // =====================================================
        private static void WriteUrdfFile(RobotModel robot, string urdfPath)
        {
            if (robot == null || string.IsNullOrEmpty(urdfPath))
                return;

            StringBuilder sb = new StringBuilder(1024 * 256);

            sb.AppendLine("<?xml version=\"1.0\"?>");
            sb.AppendLine("<robot name=\"" + XmlEscape(robot.Name ?? "robot") + "\">");

            // LINKS
            foreach (UrdfLink link in robot.Links)
            {
                if (link == null || string.IsNullOrEmpty(link.Name)) continue;

                sb.AppendLine("  <link name=\"" + XmlEscape(link.Name) + "\">");

                // inertial (si hay mesh, igual agregamos inertial básico)
                if (link.InertialMass > 0.0)
                {
                    double[] ixyz = link.InertialOriginXYZ ?? new double[] { 0, 0, 0 };

                    sb.Append("    <inertial>\n");
                    sb.Append("      <origin xyz=\"").Append(FormatXYZ(ixyz)).Append("\" rpy=\"0 0 0\"/>\n");
                    sb.Append("      <mass value=\"").Append(FormatF(link.InertialMass)).Append("\"/>\n");
                    sb.Append("      <inertia ");
                    sb.Append("ixx=\"").Append(FormatF(link.InertiaIxx)).Append("\" ");
                    sb.Append("ixy=\"").Append(FormatF(link.InertiaIxy)).Append("\" ");
                    sb.Append("ixz=\"").Append(FormatF(link.InertiaIxz)).Append("\" ");
                    sb.Append("iyy=\"").Append(FormatF(link.InertiaIyy)).Append("\" ");
                    sb.Append("iyz=\"").Append(FormatF(link.InertiaIyz)).Append("\" ");
                    sb.Append("izz=\"").Append(FormatF(link.InertiaIzz)).Append("\"/>\n");
                    sb.Append("    </inertial>\n");
                }

                // visual/collision si hay malla
                if (!string.IsNullOrEmpty(link.MeshFile))
                {
                    sb.Append("    <visual>\n");
                    sb.Append("      <origin xyz=\"0 0 0\" rpy=\"0 0 0\"/>\n");
                    sb.Append("      <geometry>\n");
                    sb.Append("        <mesh filename=\"").Append(XmlEscape(link.MeshFile)).Append("\"/>\n");
                    sb.Append("      </geometry>\n");
                    sb.Append("      <material name=\"mat_").Append(XmlEscape(link.Name)).Append("\">\n");
                    sb.Append("        <texture filename=\"meshes/").Append(XmlEscape(link.Name)).Append(".png\"/>\n");
                    sb.Append("      </material>\n");
                    sb.Append("    </visual>\n");

                    sb.Append("    <collision>\n");
                    sb.Append("      <origin xyz=\"0 0 0\" rpy=\"0 0 0\"/>\n");
                    sb.Append("      <geometry>\n");
                    sb.Append("        <mesh filename=\"").Append(XmlEscape(link.MeshFile)).Append("\"/>\n");
                    sb.Append("      </geometry>\n");
                    sb.Append("    </collision>\n");
                }

                sb.AppendLine("  </link>");
            }

            // JOINTS
            foreach (UrdfJoint j in robot.Joints)
            {
                if (j == null || string.IsNullOrEmpty(j.Name)) continue;
                if (string.IsNullOrEmpty(j.ParentLink) || string.IsNullOrEmpty(j.ChildLink)) continue;

                string type = string.IsNullOrEmpty(j.Type) ? "fixed" : j.Type;

                sb.Append("  <joint name=\"").Append(XmlEscape(j.Name)).Append("\" type=\"").Append(XmlEscape(type)).Append("\">\n");
                sb.Append("    <parent link=\"").Append(XmlEscape(j.ParentLink)).Append("\"/>\n");
                sb.Append("    <child link=\"").Append(XmlEscape(j.ChildLink)).Append("\"/>\n");

                double[] oxyz = j.OriginXYZ ?? new double[] { 0, 0, 0 };
                double[] orpy = j.OriginRPY ?? new double[] { 0, 0, 0 };
                sb.Append("    <origin xyz=\"").Append(FormatXYZ(oxyz)).Append("\" rpy=\"").Append(FormatXYZ(orpy)).Append("\"/>\n");

                if (!string.Equals(type, "fixed", StringComparison.OrdinalIgnoreCase))
                {
                    double[] axis = (j.AxisXYZ != null && j.AxisXYZ.Length == 3) ? j.AxisXYZ : new double[] { 0, 0, 1 };
                    sb.Append("    <axis xyz=\"").Append(FormatXYZ(axis)).Append("\"/>\n");

                    // limits (solo para revolute/prismatic)
                    if (string.Equals(type, "revolute", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(type, "prismatic", StringComparison.OrdinalIgnoreCase))
                    {
                        double lower = j.LimitLower;
                        double upper = j.LimitUpper;

                        // defaults suaves (evitar URDF inválido)
                        double effort = (j.LimitEffort > 0) ? j.LimitEffort : 10.0;
                        double vel = (j.LimitVelocity > 0) ? j.LimitVelocity : 1.0;

                        // Si no definiste rangos, ponemos algo razonable
                        if (string.Equals(type, "revolute", StringComparison.OrdinalIgnoreCase))
                        {
                            if (Math.Abs(upper - lower) < 1e-9)
                            {
                                lower = -Math.PI;
                                upper = Math.PI;
                            }
                        }
                        else
                        {
                            if (Math.Abs(upper - lower) < 1e-9)
                            {
                                lower = -0.1;
                                upper = 0.1;
                            }
                        }

                        sb.Append("    <limit lower=\"").Append(FormatF(lower))
                          .Append("\" upper=\"").Append(FormatF(upper))
                          .Append("\" effort=\"").Append(FormatF(effort))
                          .Append("\" velocity=\"").Append(FormatF(vel))
                          .Append("\"/>\n");
                    }
                }

                sb.Append("  </joint>\n");
            }

            sb.AppendLine("</robot>");

            IOFile.WriteAllText(urdfPath, sb.ToString(), Encoding.UTF8);
        }

        private static string FormatXYZ(double[] v)
        {
            if (v == null || v.Length < 3) return "0 0 0";
            return FormatF(v[0]) + " " + FormatF(v[1]) + " " + FormatF(v[2]);
        }

        // =====================================================
        //  Matrix → RPY (rad)  (roll X, pitch Y, yaw Z)
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
                roll = Math.Atan2(r32, r33);
                yaw = Math.Atan2(r21, r11);
            }
            else
            {
                pitch = Math.Atan2(-r31, sy);
                roll = 0.0;
                yaw = Math.Atan2(-r12, r22);
            }
        }
    }

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

        // Para compatibilidad (se usa en tu builder; la visual va en origen 0,0,0)
        public double[] OriginXYZ;
        public double[] OriginRPY;

        // malla (relativa dentro de exportDir)
        public string MeshFile;

        // Inertial (URDF)
        public double InertialMass = 1.0;
        public double[] InertialOriginXYZ = new double[] { 0, 0, 0 };

        public double InertiaIxx = 1e-4;
        public double InertiaIxy = 0.0;
        public double InertiaIxz = 0.0;
        public double InertiaIyy = 1e-4;
        public double InertiaIyz = 0.0;
        public double InertiaIzz = 1e-4;
    }

    public class UrdfJoint
    {
        public string Name;
        public string Type; // fixed / revolute / continuous / prismatic

        public string ParentLink;
        public string ChildLink;

        public double[] OriginXYZ = new double[] { 0, 0, 0 };
        public double[] OriginRPY = new double[] { 0, 0, 0 };

        // axis en el frame del joint (normalmente frame del child/origin)
        public double[] AxisXYZ = null;

        // Limits (solo para revolute/prismatic)
        public double LimitLower = 0.0;
        public double LimitUpper = 0.0;
        public double LimitEffort = 0.0;
        public double LimitVelocity = 0.0;
    }
}

