using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using Inventor;

namespace URDFConverterAddIn
{
    // GUID NUEVO: 4aba8fec-e1bf-4786-ba37-6163b7bc0953
    [Guid("4aba8fec-e1bf-4786-ba37-6163b7bc0953")]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application _invApp;

        // Dos botones: VeryLowOptimized y DisplayMesh
        private ButtonDefinition _exportUrdfVlqButton;
        private ButtonDefinition _exportUrdfDisplayButton;

        // ----------------------------------------------------
        //  Activate: se ejecuta cuando Inventor carga el AddIn
        //  (agrega panel en Part y en Assembly)
        // ----------------------------------------------------
        public void Activate(ApplicationAddInSite AddInSiteObject, bool FirstTime)
        {
            _invApp = AddInSiteObject.Application;

            try
            {
                CommandManager cmdMgr = _invApp.CommandManager;

                // Botón 1: Very Low Quality Optimized (VLQ)
                _exportUrdfVlqButton = cmdMgr.ControlDefinitions.AddButtonDefinition(
                    "Export URDF (VLQ)",            // DisplayName
                    "urdf_export_vlq_cmd",          // InternalName (único)
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "4aba8fece1bf4786ba376163b7bc0953", // ClientId sin guiones
                    "Export URDF with VeryLowOptimized mesh",
                    "Export URDF (VLQ)");

                _exportUrdfVlqButton.OnExecute +=
                    new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfVlqButtonPressed);

                // Botón 2: DisplayMesh (alta calidad)
                _exportUrdfDisplayButton = cmdMgr.ControlDefinitions.AddButtonDefinition(
                    "Export URDF (Display)",        // DisplayName
                    "urdf_export_display_cmd",      // InternalName (único)
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "5017703b3b0d4c6ea5590ae90e268c2f", // mismo ClientId
                    "Export URDF with DisplayMesh-quality mesh",
                    "Export URDF (Display)");

                _exportUrdfDisplayButton.OnExecute +=
                    new ButtonDefinitionSink_OnExecuteEventHandler(OnExportUrdfDisplayButtonPressed);

                // -------------------------------------------------
                //  Añadir los botones a los Ribbons de Part y Assembly
                // -------------------------------------------------
                UserInterfaceManager uiMgr = _invApp.UserInterfaceManager;

                // 1) Ribbon de PIEZAS (Part)
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
                            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                            "",
                            false);
                    }

                    urdfPanelPart.CommandControls.AddButton(_exportUrdfVlqButton);
                    urdfPanelPart.CommandControls.AddButton(_exportUrdfDisplayButton);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[URDF][UI] Error creando panel en Part: " + ex.Message);
                }

                // 2) Ribbon de ENSAMBLAJES (Assembly)
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
                            "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
                            "",
                            false);
                    }

                    urdfPanelAsm.CommandControls.AddButton(_exportUrdfVlqButton);
                    urdfPanelAsm.CommandControls.AddButton(_exportUrdfDisplayButton);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[URDF][UI] Error creando panel en Assembly: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error activating URDFConverter AddIn:\n" + ex.Message,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------
        //  OnExportUrdfVlqButtonPressed
        //  - Llama al exportador con VeryLowOptimized
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
        //  OnExportUrdfDisplayButtonPressed
        //  - Llama al exportador con DisplayMesh
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

        // ----------------------------------------------------
        //  Deactivate: liberar referencias COM
        // ----------------------------------------------------
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
                // ignorar
            }

            if (_invApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_invApp);
                }
                catch
                {
                }
                _invApp = null;
            }
        }

        // ----------------------------------------------------
        //  ExecuteCommand: no lo usamos, pero hay que implementarlo
        // ----------------------------------------------------
        public void ExecuteCommand(int CommandID)
        {
            // No implementado (requerido por la interfaz)
        }

        // ----------------------------------------------------
        //  Automation: para exponer objetos COM (no lo usamos)
        // ----------------------------------------------------
        public object Automation
        {
            get { return null; }
        }
    }

    public static class UrdfExporter
    {
        // -------------------------------------------------
        //  MODO DE CALIDAD DE MALLA
        // -------------------------------------------------
        private static string _meshQualityMode = "very_low_optimized";

        public static void SetMeshQualityVeryLow()
        {
            _meshQualityMode = "very_low_optimized";
        }

        public static void SetMeshQualityDisplay()
        {
            _meshQualityMode = "display_mesh";
        }

        public static string GetMeshQualityMode()
        {
            return _meshQualityMode;
        }

        // -------------------------------------------------
        //  FLAGS DE DEBUG
        // -------------------------------------------------
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
            Trace.WriteLine(full);   // para DebugView
        }

        // =====================================================
        //  ExportActiveDocument
        // =====================================================
        public static void ExportActiveDocument(Inventor.Application invApp)
        {
            if (invApp == null)
            {
                MessageBox.Show(
                    "Inventor.Application es nulo.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            Document doc = invApp.ActiveDocument as Document;
            if (doc == null)
            {
                MessageBox.Show(
                    "No hay documento activo para exportar.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Solo soportamos por ahora Part y Assembly
            if (doc.DocumentType != DocumentTypeEnum.kPartDocumentObject &&
                doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
            {
                MessageBox.Show(
                    "Solo se soportan documentos de pieza (.ipt) y ensamblaje (.iam).",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string fullPath = string.Empty;
            try
            {
                fullPath = doc.FullFileName;
            }
            catch
            {
                fullPath = string.Empty;
            }

            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show(
                    "El documento no tiene ruta de fichero. Guárdalo antes de exportar.",
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string baseDir = IOPath.GetDirectoryName(fullPath);
            string baseName = IOPath.GetFileNameWithoutExtension(fullPath);

            DebugLog(
                "SYS",
                "ExportActiveDocument: doc='" + doc.DisplayName + "', type=" +
                doc.DocumentType.ToString() + ", path='" + fullPath +
                "', meshMode=" + _meshQualityMode);

            string exportDir = IOPath.Combine(baseDir, "URDF_Export");
            if (!EnsureDirectory(exportDir))
            {
                MessageBox.Show(
                    "No se pudo crear la carpeta de exportación:\n" + exportDir,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            string urdfPath = IOPath.Combine(exportDir, baseName + ".urdf");

            try
            {
                // 1) Construir el modelo URDF en memoria (links/joints)
                RobotModel robot = BuildRobotFromDocument(doc, exportDir, baseName);

                // 2) Exportar geometría y rellenar MeshFile de cada link
                ExportGeometryAndFillMeshFiles(invApp, doc, robot, exportDir);

                // 3) Escribir el .urdf a disco
                WriteUrdfFile(robot, urdfPath);

                DebugLog("SYS", "URDF escrito en: " + urdfPath);

                MessageBox.Show(
                    "Exportación URDF completada:\n" + urdfPath,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    "Error durante la exportación URDF:\n" + ex.Message,
                    "URDFConverterAddIn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // -------------------------------------------------
        //  EnsureDirectory: equivalente a ensure_dir de Python
        // -------------------------------------------------
        private static bool EnsureDirectory(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            try
            {
                if (Directory.Exists(path))
                    return true;

                Directory.CreateDirectory(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // -------------------------------------------------
        //  BuildRobotFromDocument:
        //  Construye la estructura base_link + links hijos
        //  *** ADAPTADO A: UN LINK POR SURFACEBODY ***
        // -------------------------------------------------
        private static RobotModel BuildRobotFromDocument(
            Document doc,
            string exportDir,
            string baseName)
        {
            RobotModel robot = new RobotModel();
            robot.Name = baseName;

            DebugLog("SYS", "BuildRobotFromDocument: type=" + doc.DocumentType.ToString());

            // Link raíz base_link
            UrdfLink baseLink = new UrdfLink();
            baseLink.Name = "base_link";
            baseLink.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
            baseLink.OriginRPY = new double[] { 0.0, 0.0, 0.0 };
            robot.Links.Add(baseLink);

            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument partDoc = (PartDocument)doc;
                AddPartBodiesAsLinks(partDoc, robot, baseName);
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = (AssemblyDocument)doc;
                AddAssemblyOccurrencesAndBodiesAsLinks(asmDoc, robot);
            }

            DebugLog("SYS",
                "Robot construido: links=" + robot.Links.Count +
                ", joints=" + robot.Joints.Count);

            return robot;
        }

        // -------------------------------------------------
        //  AddPartBodiesAsLinks
        //  *** UN LINK POR SURFACEBODY EN ROOT (como root_body_i...) ***
        // -------------------------------------------------
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

            // Si no hay cuerpos, creamos un solo link genérico
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

            // Un link por cuerpo, todos colgando de base_link (root_body_i_...)
            for (int i = 0; i < bodies.Count; i++)
            {
                SurfaceBody b = bodies[i];
                string bodyName = "(null)";
                try
                {
                    if (b != null && !string.IsNullOrEmpty(b.Name))
                        bodyName = b.Name;
                }
                catch
                {
                }

                string linkName = "root_body_" +
                                  i.ToString(CultureInfo.InvariantCulture) + "_" +
                                  MakeSafeName(bodyName);

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
                    "AddPartBodiesAsLinks: creado link '" + linkName +
                    "' para SurfaceBody[" + i.ToString(CultureInfo.InvariantCulture) + "]");
            }
        }














        // -------------------------------------------------
        //  AddAssemblyOccurrencesAndBodiesAsLinks (AllLeaf)
        //  *** UN LINK POR BODY, RESPETANDO OCCURRENCES ***
        //  *** NOMBRES ÚNICOS: link_<occIndex>_<occName>[_bN] ***
        // -------------------------------------------------
        private static void AddAssemblyOccurrencesAndBodiesAsLinks(
            AssemblyDocument asmDoc,
            RobotModel robot)
        {
            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs = asmDef.Occurrences;

            // Todas las occurrences hoja
            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            // Inventor en cm → URDF en m
            double scaleToMeters = 0.01;

            DebugLog("SYS",
                "AddAssemblyOccurrencesAndBodiesAsLinks: leafOccs=" + leafOccs.Count);

            int occIndex = 0; // índice global para diferenciar occurrences con el mismo Name

            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    // Saltar componentes suprimidos / ocultos
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

                    // Recoger todos los SurfaceBody de esta occurrence
                    List<SurfaceBody> bodies = new List<SurfaceBody>();
                    CollectSurfaceBodiesFromOccurrence(occ, bodies);

                    if (bodies.Count == 0)
                    {
                        DebugLog("MESH",
                            "occ '" + occ.Name +
                            "': sin SurfaceBodies/WorkSurfaces para exportar.");
                        continue;
                    }

                    // Transformación absoluta de la occurrence
                    Matrix m = occ.Transformation;

                    double tx_m = m.Cell[1, 4] * scaleToMeters;
                    double ty_m = m.Cell[2, 4] * scaleToMeters;
                    double tz_m = m.Cell[3, 4] * scaleToMeters;

                    double roll, pitch, yaw;
                    MatrixToRPY(m, out roll, out pitch, out yaw);

                    DebugLog(
                        "TFM",
                        "occ='" + occ.Name + "' T_world(m)=(" +
                        tx_m.ToString(CultureInfo.InvariantCulture) + ", " +
                        ty_m.ToString(CultureInfo.InvariantCulture) + ", " +
                        tz_m.ToString(CultureInfo.InvariantCulture) + ") " +
                        "rpy(rad)=(" +
                        roll.ToString(CultureInfo.InvariantCulture) + ", " +
                        pitch.ToString(CultureInfo.InvariantCulture) + ", " +
                        yaw.ToString(CultureInfo.InvariantCulture) + ")"
                    );

                    string rawName = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    // Nombre base único por occurrence
                    string baseLinkName = "link_" +
                                          occIndex.ToString(CultureInfo.InvariantCulture) +
                                          "_" + safeName;

                    // Para cada SurfaceBody de la occurrence, un link:
                    //  - i == 0 → link_<idx>_<occ>
                    //  - i >= 1 → link_<idx>_<occ>_b1, _b2, ...
                    for (int i = 0; i < bodies.Count; i++)
                    {
                        SurfaceBody b = bodies[i];

                        string suffix = (i == 0)
                            ? ""
                            : "_b" + i.ToString(CultureInfo.InvariantCulture);

                        string linkName = baseLinkName + suffix;

                        // YA NO comprobamos duplicados por nombre:
                        // cada occurrence hoja tiene su propio índice global
                        UrdfLink link = new UrdfLink();
                        link.Name = linkName;
                        link.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                        link.OriginRPY = new double[] { roll, pitch, yaw };
                        robot.Links.Add(link);

                        UrdfJoint joint = new UrdfJoint();
                        joint.Type = "fixed";

                        if (i == 0)
                        {
                            // link principal cuelga de base_link (root_<name>)
                            joint.Name = "root_" + linkName;
                            joint.ParentLink = "base_link";
                            joint.ChildLink = linkName;
                            joint.OriginXYZ = new double[] { tx_m, ty_m, tz_m };
                            joint.OriginRPY = new double[] { roll, pitch, yaw };

                            DebugLog(
                                "LINK",
                                "Añadido link principal '" + linkName +
                                "' colgando de base_link.");
                        }
                        else
                        {
                            // extras cuelgan del principal con origen 0
                            joint.Name = "fixed_extra_" + linkName;
                            joint.ParentLink = baseLinkName;
                            joint.ChildLink = linkName;
                            joint.OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
                            joint.OriginRPY = new double[] { 0.0, 0.0, 0.0 };

                            DebugLog(
                                "LINK",
                                "Añadido link extra '" + linkName +
                                "' colgando de '" + baseLinkName + "'.");
                        }

                        robot.Joints.Add(joint);
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLog(
                        "ERR",
                        "Error al crear links/joints para occurrence '" +
                        occ.Name + "': " + ex.Message
                    );
                }
                finally
                {
                    // Muy importante: avanzar el índice global SIEMPRE,
                    // incluso si hicimos continue o hubo excepción.
                    occIndex++;
                }
            }
        }

        // -------------------------------------------------
        //  ExportGeometryAndFillMeshFiles (DAE)
        // -------------------------------------------------
        private static void ExportGeometryAndFillMeshFiles(
            Inventor.Application invApp,
            Document doc,
            RobotModel robot,
            string exportDir)
        {
            string meshesDir = IOPath.Combine(exportDir, "meshes");
            EnsureDirectory(meshesDir);

            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                PartDocument partDoc = (PartDocument)doc;
                ExportPartGeometryToDae(partDoc, robot, meshesDir);
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                AssemblyDocument asmDoc = (AssemblyDocument)doc;
                ExportAssemblyGeometryToDae(asmDoc, robot, meshesDir);
            }
        }

        // -------------------------------------------------
        //  Helpers: recoger SurfaceBody para exportar
        // -------------------------------------------------
        private static void CollectSurfaceBodiesFromPartDefinition(
            PartComponentDefinition partDef,
            List<SurfaceBody> bodies)
        {
            if (partDef == null || bodies == null)
                return;

            // 1) SurfaceBodies "normales"
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
            catch
            {
                // ignoramos errores puntuales
            }

            // 2) Sheet / WorkSurfaces → SurfaceBodies
            try
            {
                WorkSurfaces workSurfaces = partDef.WorkSurfaces;
                if (workSurfaces != null)
                {
                    for (int wi = 1; wi <= workSurfaces.Count; wi++)
                    {
                        WorkSurface ws = workSurfaces[wi];
                        if (ws == null)
                            continue;

                        SurfaceBodies wsBodies = ws.SurfaceBodies;
                        if (wsBodies == null)
                            continue;

                        for (int bi = 1; bi <= wsBodies.Count; bi++)
                        {
                            SurfaceBody b = wsBodies[bi];
                            if (b != null)
                                bodies.Add(b);
                        }
                    }
                }
            }
            catch
            {
                // ignoramos errores puntuales
            }
        }

        // *** MÉTODO CORREGIDO: evita duplicar cuerpos ***
        private static void CollectSurfaceBodiesFromOccurrence(
            ComponentOccurrence occ,
            List<SurfaceBody> bodies)
        {
            if (occ == null || bodies == null)
                return;

            // 1) Intentar usar SOLO los cuerpos en contexto de ensamblaje
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

                    // Si ya hay cuerpos proxy, NO miramos el PartDefinition
                    return;
                }
            }
            catch
            {
                // ignoramos
            }

            // 2) Solo si no hay cuerpos proxy, usamos el PartComponentDefinition (fallback)
            try
            {
                PartComponentDefinition partDef = occ.Definition as PartComponentDefinition;
                if (partDef != null)
                    CollectSurfaceBodiesFromPartDefinition(partDef, bodies);
            }
            catch
            {
                // ignoramos
            }
        }

        // -------------------------------------------------
        //  ExportPartGeometryToDae  (PIEZA .ipt)
        //  *** UN DAE POR BODY: root_body_i_* ***
        // -------------------------------------------------
        private static void ExportPartGeometryToDae(
            PartDocument partDoc,
            RobotModel robot,
            string meshesDir)
        {
            string baseName = IOPath.GetFileNameWithoutExtension(partDoc.DisplayName);

            PartComponentDefinition partDef = partDoc.ComponentDefinition;

            // Recogemos TODOS los SurfaceBody relevantes
            List<SurfaceBody> bodies = new List<SurfaceBody>();
            CollectSurfaceBodiesFromPartDefinition(partDef, bodies);

            DebugLog("MESH",
                "ExportPartGeometryToDae: Part '" + baseName +
                "': SurfaceBodies recogidos=" + bodies.Count);

            if (bodies.Count == 0)
            {
                DebugLog("MESH",
                    "Part '" + baseName + "': sin SurfaceBodies para exportar.");
                return;
            }

            for (int i = 0; i < bodies.Count; i++)
            {
                SurfaceBody body = bodies[i];
                if (body == null)
                    continue;

                string bodyName = "(null)";
                try
                {
                    if (!string.IsNullOrEmpty(body.Name))
                        bodyName = body.Name;
                }
                catch
                {
                }

                string linkName = "root_body_" +
                                  i.ToString(CultureInfo.InvariantCulture) + "_" +
                                  MakeSafeName(bodyName);

                UrdfLink link = FindLinkByName(robot, linkName);
                if (link == null)
                {
                    DebugLog("MESH",
                        "ExportPartGeometryToDae: no se encontró link '" +
                        linkName + "' para body[" + i.ToString(CultureInfo.InvariantCulture) + "]");
                    continue;
                }

                // Tessellate SOLO este body
                List<SurfaceBody> oneBodyList = new List<SurfaceBody>();
                oneBodyList.Add(body);

                double[] vertices;
                int[] indices;
                if (!TessellateBodiesToMeshArrays(oneBodyList, out vertices, out indices))
                {
                    DebugLog("MESH",
                        "Part '" + baseName + "', body[" +
                        i.ToString(CultureInfo.InvariantCulture) +
                        "]: tessellate no generó triángulos.");
                    continue;
                }

                string fileName = linkName + ".dae";
                string fullPath = IOPath.Combine(meshesDir, fileName);

                // Para parts, los vértices ya están en coords locales del body (origen del .ipt)
                WriteColladaFile(fullPath, linkName, vertices, indices);
                DebugLog("SYS",
                    "Part '" + baseName + "', body[" +
                    i.ToString(CultureInfo.InvariantCulture) + "]: DAE '" +
                    fileName + "' (verts=" + (vertices.Length / 3) +
                    ", tris=" + (indices.Length / 3) + ")");

                // Ruta RELATIVA desde el .urdf
                link.MeshFile = "meshes/" + fileName;

                // ---- INERCIAL desde MassProperties del Part (global) ----
                try
                {
                    MassProperties mp = partDef.MassProperties;
                    FillLinkInertialFromMassProperties(link, mp);
                }
                catch
                {
                    // Si falla, el link se queda con inercial dummy
                }
            }
        }

        // -------------------------------------------------
        //  ExportAssemblyGeometryToDae  (ENSAMBLADO .iam)
        //  *** UN DAE POR BODY: link_<idx>_<occ>, link_<idx>_<occ>_b1... ***
        // -------------------------------------------------
        private static void ExportAssemblyGeometryToDae(
            AssemblyDocument asmDoc,
            RobotModel robot,
            string meshesDir)
        {
            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
            ComponentOccurrences occs = asmDef.Occurrences;

            ComponentOccurrencesEnumerator leafOccs = occs.AllLeafOccurrences;

            int occIndex = 0; // Debe evolucionar igual que en AddAssemblyOccurrencesAndBodiesAsLinks

            foreach (ComponentOccurrence occ in leafOccs)
            {
                try
                {
                    // Saltar componentes suprimidos / ocultos
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

                    string rawName = occ.Name;
                    string safeName = MakeSafeName(rawName);

                    // Debe coincidir EXACTAMENTE con AddAssemblyOccurrencesAndBodiesAsLinks
                    string baseLinkName = "link_" +
                                          occIndex.ToString(CultureInfo.InvariantCulture) +
                                          "_" + safeName;

                    // Recoger todos los SurfaceBody de esta occurrence
                    List<SurfaceBody> bodies = new List<SurfaceBody>();
                    CollectSurfaceBodiesFromOccurrence(occ, bodies);

                    DebugLog("MESH",
                        "occ '" + rawName + "': SurfaceBodies recogidos=" + bodies.Count);

                    if (bodies.Count == 0)
                    {
                        DebugLog("MESH",
                            "occ '" + rawName +
                            "': sin SurfaceBodies/WorkSurfaces para exportar.");
                        continue;
                    }

                    // Matriz de transformación de la occurrence (para pasar a coords locales)
                    Matrix m = occ.Transformation;

                    for (int i = 0; i < bodies.Count; i++)
                    {
                        SurfaceBody body = bodies[i];
                        if (body == null)
                            continue;

                        string suffix = (i == 0)
                            ? ""
                            : "_b" + i.ToString(CultureInfo.InvariantCulture);

                        string linkName = baseLinkName + suffix;

                        UrdfLink link = FindLinkByName(robot, linkName);
                        if (link == null)
                        {
                            DebugLog("MESH",
                                "occ '" + rawName + "', body[" +
                                i.ToString(CultureInfo.InvariantCulture) +
                                "]: no hay link '" + linkName + "', se omite.");
                            continue;
                        }

                        // Tessellate SOLO este body (vértices inicialmente en coords world de la occ)
                        double[] verticesWorld;
                        int[] indices;

                        List<SurfaceBody> oneBodyList = new List<SurfaceBody>();
                        oneBodyList.Add(body);

                        if (!TessellateBodiesToMeshArrays(oneBodyList, out verticesWorld, out indices))
                        {
                            DebugLog("MESH",
                                "occ '" + rawName + "', body[" +
                                i.ToString(CultureInfo.InvariantCulture) +
                                "]: tessellate no generó triángulos.");
                            continue;
                        }

                        // Convertir vértices world (m) → coords LOCALES del componente
                        double[] verticesLocal;
                        TransformVerticesToLocalFrame(verticesWorld, m, out verticesLocal);

                        // Nombre DAE = nombre del link (único por diseño)
                        string fileName = linkName + ".dae";
                        string fullPath = IOPath.Combine(meshesDir, fileName);

                        WriteColladaFile(fullPath, linkName, verticesLocal, indices);

                        // Ruta relativa (desde la carpeta del .urdf)
                        link.MeshFile = "meshes/" + fileName;

                        // ---- INERCIAL desde MassProperties del PartDefinition ----
                        try
                        {
                            PartComponentDefinition partDef = occ.Definition as PartComponentDefinition;
                            if (partDef != null)
                            {
                                MassProperties mp = partDef.MassProperties;
                                FillLinkInertialFromMassProperties(link, mp);
                            }
                        }
                        catch
                        {
                            // Si falla, seguimos; el link tendrá inercial dummy
                        }

                        DebugLog("MESH",
                            "occ '" + rawName + "', body[" +
                            i.ToString(CultureInfo.InvariantCulture) + "]: escribió " +
                            fileName + " con " + (verticesLocal.Length / 3) +
                            " vértices y " + (indices.Length / 3) + " triángulos.");
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLog("ERR",
                        "Error al exportar geometría para occ '" +
                        occ.Name + "': " + ex.Message);
                }
                finally
                {
                    // Igual que en AddAssemblyOccurrencesAndBodiesAsLinks
                    occIndex++;
                }
            }
        }




















        // -------------------------------------------------
        //  FindLinkByName
        // -------------------------------------------------
        private static UrdfLink FindLinkByName(RobotModel robot, string name)
        {
            if (robot == null || robot.Links == null)
                return null;

            for (int i = 0; i < robot.Links.Count; i++)
            {
                UrdfLink link = robot.Links[i];
                if (link != null && link.Name == name)
                    return link;
            }

            return null;
        }

        // -------------------------------------------------
        //  TessellateBodiesToMeshArrays
        //  (usa el sistema CalculateFacets del DAE converter bueno)
        //  *** SE USA PARA LISTAS, PERO AQUÍ SIEMPRE PASAMOS 1 BODY ***
        // -------------------------------------------------
        private static bool TessellateBodiesToMeshArrays(
            IList<SurfaceBody> bodies,
            out double[] vertices,
            out int[] indices)
        {
            vertices = null;
            indices  = null;

            if (bodies == null || bodies.Count == 0)
                return false;

            List<double> vList = new List<double>();
            List<int>    iList = new List<int>();

            int vertexOffset = 0;

            foreach (SurfaceBody body in bodies)
            {
                if (body == null)
                    continue;

                if (!TessellateSingleBody(body, vList, iList, ref vertexOffset))
                {
                    DebugLog("MESH", "Body sin triángulos (CalculateFacets), se omite.");
                }
            }

            if (vList.Count == 0 || iList.Count == 0)
                return false;

            vertices = vList.ToArray();
            indices  = iList.ToArray();
            return true;
        }

        // -------------------------------------------------
        //  TessellateSingleBody - patrón del DAE converter que funciona
        // -------------------------------------------------
        private static bool TessellateSingleBody(
            SurfaceBody body,
            List<double> vList,
            List<int> iList,
            ref int vertexOffset)
        {
            try
            {
                // Tolerancia en cm (API Inventor → cm)
                double tol;
                if (_meshQualityMode == "display_mesh")
                    tol = 0.05;   // más fino
                else
                    tol = 0.1;    // modo rápido

                int vertexCount  = 0;
                int facetCount   = 0;

                // IMPORTANTE: inicializar arrays vacíos (igual que DAE converter)
                double[] vertexCoords   = new double[] { };
                double[] normalVectors  = new double[] { };
                int[]    vertexIndices  = new int[] { };

                body.CalculateFacets(
                    tol,
                    out vertexCount,
                    out facetCount,
                    out vertexCoords,
                    out normalVectors,
                    out vertexIndices);

                if (vertexCount <= 0 || facetCount <= 0 ||
                    vertexCoords == null || vertexCoords.Length == 0 ||
                    vertexIndices == null || vertexIndices.Length == 0)
                {
                    DebugLog("MESH",
                        "CalculateFacets devolvió 0 vértices o 0 facetas.");
                    return false;
                }

                // Convertir vértices cm → m
                for (int i = 0; i < vertexCoords.Length; i++)
                {
                    double vCm = vertexCoords[i];
                    double vMeters = vCm * 0.01;
                    vList.Add(vMeters);
                }

                // Índices 1-based → 0-based con offset
                for (int i = 0; i < vertexIndices.Length; i++)
                {
                    int idx = vertexIndices[i] - 1;
                    if (idx < 0) idx = 0;
                    iList.Add(vertexOffset + idx);
                }

                vertexOffset = vList.Count / 3;

                DebugLog("MESH",
                    "Body tessellated: verts=" + vertexCount +
                    ", facets=" + facetCount +
                    ", newOffset=" + vertexOffset);

                return true;
            }
            catch (Exception ex)
            {
                DebugLog("ERR", "Error en TessellateSingleBody: " + ex.Message);
                return false;
            }
        }

        // -------------------------------------------------
        //  TransformVerticesToLocalFrame
        //  Convierte vértices en coords WORLD (m) → coords LOCALES del componente
        // -------------------------------------------------
        private static void TransformVerticesToLocalFrame(
            double[] verticesWorld,
            Matrix occMatrix,
            out double[] verticesLocal)
        {
            verticesLocal = null;
            if (verticesWorld == null || verticesWorld.Length == 0 || occMatrix == null)
            {
                verticesLocal = verticesWorld;
                return;
            }

            double scaleToMeters = 0.01;

            // Traslación de la occurrence en metros
            double tx = occMatrix.Cell[1, 4] * scaleToMeters;
            double ty = occMatrix.Cell[2, 4] * scaleToMeters;
            double tz = occMatrix.Cell[3, 4] * scaleToMeters;

            // Matriz de rotación R
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
        }

        // -------------------------------------------------
        //  FillLinkInertialFromMassProperties
        //  (COM local: COM_global - OriginXYZ del link)
        // -------------------------------------------------
        private static void FillLinkInertialFromMassProperties(
            UrdfLink link,
            MassProperties mp)
        {
            if (link == null || mp == null)
                return;

            double mass = mp.Mass;

            // Centro de masa (en cm → pasamos a m, GLOBAL)
            Point com = mp.CenterOfMass;
            double scaleToMeters = 0.01;
            double comGlobalX = com.X * scaleToMeters;
            double comGlobalY = com.Y * scaleToMeters;
            double comGlobalZ = com.Z * scaleToMeters;

            // COM LOCAL al frame del link
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

            double Ixx;
            double Iyy;
            double Izz;
            double Ixy;
            double Iyz;
            double Ixz;

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
                "kg, COM_local=(" +
                comLocalX.ToString("G5", CultureInfo.InvariantCulture) + "," +
                comLocalY.ToString("G5", CultureInfo.InvariantCulture) + "," +
                comLocalZ.ToString("G5", CultureInfo.InvariantCulture) + ")");
        }

        // -------------------------------------------------
        //  WriteColladaFile: escribe un DAE mínimo
        // -------------------------------------------------
        private static void WriteColladaFile(
            string fullPath,
            string geometryName,
            double[] vertices,
            int[] indices)
        {
            string text = BuildColladaText(geometryName, vertices, indices);
            IOFile.WriteAllText(fullPath, text);
        }

        private static string BuildColladaText(
            string geometryName,
            double[] vertices,
            int[] indices)
        {
            StringBuilder sb = new StringBuilder();

            string geomId = geometryName + "-geom";
            string positionsId = geometryName + "-positions";
            string positionsArrayId = positionsId + "-array";
            string verticesId = geometryName + "-verts";

            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine("<COLLADA xmlns=\"http://www.collada.org/2005/11/COLLADASchema\" version=\"1.4.1\">");
            sb.AppendLine("  <asset>");
            sb.AppendLine("    <contributor>");
            sb.AppendLine("      <authoring_tool>URDFConverterAddIn</authoring_tool>");
            sb.AppendLine("    </contributor>");
            sb.AppendLine("    <unit name=\"meter\" meter=\"1\"/>");
            sb.AppendLine("    <up_axis>Z_UP</up_axis>");
            sb.AppendLine("  </asset>");

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

            // Triangles
            int triCount = indices.Length / 3;
            sb.AppendLine("        <triangles count=\"" + triCount.ToString(CultureInfo.InvariantCulture) + "\">");
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
            sb.AppendLine("        <instance_geometry url=\"#" + geomId + "\"/>");
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

        // -------------------------------------------------
        //  WriteUrdfFile: genera el .urdf completo
        // -------------------------------------------------
        private static void WriteUrdfFile(RobotModel robot, string urdfPath)
        {
            StringBuilder sb = new StringBuilder();

            string robotName = robot.Name;
            if (string.IsNullOrEmpty(robotName))
                robotName = "InventorRobot";

            sb.AppendLine("<?xml version=\"1.0\"?>");
            sb.AppendLine("<robot name=\"" + XmlEscape(robotName) + "\">");

            // Links
            foreach (UrdfLink link in robot.Links)
            {
                if (link == null)
                    continue;

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

            // Joints
            foreach (UrdfJoint joint in robot.Joints)
            {
                if (joint == null)
                    continue;

                sb.AppendLine("  <joint name=\"" + XmlEscape(joint.Name) + "\" type=\"" + XmlEscape(joint.Type) + "\">");
                sb.AppendLine("    <parent link=\"" + XmlEscape(joint.ParentLink) + "\"/>");
                sb.AppendLine("    <child link=\"" + XmlEscape(joint.ChildLink) + "\"/>");

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
                sb.AppendLine("  </joint>");
            }

            sb.AppendLine("</robot>");

            IOFile.WriteAllText(urdfPath, sb.ToString());
        }

        // -------------------------------------------------
        //  XmlEscape: escapa &, <, >, "
        // -------------------------------------------------
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

        // -------------------------------------------------
        //  MatrixToRPY: matriz de rotación → roll-pitch-yaw
        // -------------------------------------------------
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
        }

        // -------------------------------------------------
        //  MakeSafeName: limpia nombres para URDF
        // -------------------------------------------------
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
    }

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
            Links = new List<UrdfLink>();
            Joints = new List<UrdfJoint>();
        }
    }

    // -------------------------------------------------
    //  URDF LINK (con datos de inercia)
    // -------------------------------------------------
    public class UrdfLink
    {
        public string Name;

        // Geometría visual (archivo DAE relativo al .urdf)
        public string MeshFile;

        // Pose del link respecto a base_link (o parent) en URDF
        public double[] OriginXYZ;   // [x, y, z] en metros
        public double[] OriginRPY;   // [roll, pitch, yaw] en rad

        // ---- INERCIAL (opcional) ----
        public bool HasInertial = false;

        // Masa en kg
        public double Mass = 1e-6;

        // Origen de la inercia (COM) relativo al frame del link
        public double[] InertialOriginXYZ = new double[] { 0.0, 0.0, 0.0 };
        public double[] InertialOriginRPY = new double[] { 0.0, 0.0, 0.0 };

        // Componentes de la matriz de inercia alrededor del COM
        public double Ixx = 1e-6;
        public double Iyy = 1e-6;
        public double Izz = 1e-6;
        public double Ixy = 0.0;
        public double Iyz = 0.0;
        public double Ixz = 0.0;
    }

    public class UrdfJoint
    {
        public string Name;
        public string Type;         // "fixed", "revolute", etc.
        public string ParentLink;
        public string ChildLink;

        public double[] OriginXYZ;
        public double[] OriginRPY;

        public UrdfJoint()
        {
            Type = "fixed";
            OriginXYZ = new double[] { 0.0, 0.0, 0.0 };
            OriginRPY = new double[] { 0.0, 0.0, 0.0 };
        }
    }
}



