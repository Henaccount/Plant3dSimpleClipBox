using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Autodesk.ProcessPower.ProjectManager;

[assembly: CommandClass(typeof(ClipBox.Program))]

//commands: selectlast, clipping, prepareBox, loadRequiredXrefs <path>


namespace ClipBox
{
    public class Program
    {
        //[CommandMethod("DetachAllXrefs")]
        public static void detachallxrefs()
        {
            Database db = Helper.oDatabase;
            Editor ed = Helper.oEditor;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
                foreach (ObjectId id in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)id.GetObject(OpenMode.ForRead);
                    if (btr.XrefStatus != XrefStatus.NotAnXref)
                    {
                        db.DetachXref(id);
                    }
                }
                tr.Commit();
            }
        }

        [CommandMethod("selectLast")]
        public static void selectLast()
        {
            Helper.Initialize();
            PromptSelectionResult selRes = Helper.oEditor.SelectLast();
            Helper.oEditor.SetImpliedSelection(selRes.Value.GetObjectIds());
            Helper.Terminate();
        }

        public static Extents3d getBoxExtends()
        {
            Extents3d thext = new Extents3d();
            try
            {

                using (Transaction tr = Helper.oDatabase.TransactionManager.StartTransaction())
                {

                    TypedValue[] filterlist = new TypedValue[2];
                    filterlist[0] = new TypedValue(0, "3DSOLID");
                    filterlist[1] = new TypedValue(8, "clipbox");
                    SelectionFilter filter = new SelectionFilter(filterlist);
                    PromptSelectionResult selRes = Helper.oEditor.SelectAll(filter);
                    ObjectId[] objIdArray = selRes.Value.GetObjectIds();

                    if (objIdArray.Length == 0)
                    {
                        createBox(new Point3d(1,1,1));
                    }
                    else if (objIdArray.Length > 1)
                    {
                        Helper.oEditor.WriteMessage("\ntake action: Just one object (box) is allowed to be on the clipbox layer, returning micro box at origin..");
                        thext = new Extents3d(new Point3d(0, 0, 0), new Point3d(1, 1, 1));
                        return thext;
                    }

                    Entity ent = tr.GetObject(objIdArray[0], OpenMode.ForRead) as Entity;
                    thext = ent.Bounds.Value;
                    tr.Commit();
                }
            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage(e.Message);

            }
            finally { }

            return thext;
        }

        [CommandMethod("clipping", CommandFlags.UsePickSet)]
        public static void clipping()
        {
            try
            {

                object snapmodesaved = Application.GetSystemVariable("SNAPMODE");
                object osmodesaved = Application.GetSystemVariable("OSMODE");
                object os3dmodesaved = Application.GetSystemVariable("3DOSMODE");


                Application.SetSystemVariable("XCLIPFRAME", 0);
                Application.SetSystemVariable("SNAPMODE", 0);
                Application.SetSystemVariable("OSMODE", 0);
                Application.SetSystemVariable("3DOSMODE", 0);

                Helper.Initialize();
                Helper.oEditor.Command("_.layerclose");
                verifyLayerIsPresent();
                Helper.oEditor.Command("_Plantshowall");

                Extents3d thext = getBoxExtends();

                Point3d themax = thext.MaxPoint;
                Point3d themin = thext.MinPoint;


                Helper.oEditor.Command("_XCLIP", "_ALL", "", "_N", "_Y", "_R", themin, themax);
                Helper.oEditor.Command("_XCLIP", "_ALL", "", "_C", themax, themin);

                if (!ShowHideBox(true))
                    ShowHideBox(false);

                //now clip drawing parts
                List<ObjectId> forSelection = new List<ObjectId>();

                using (Transaction tr = Helper.oDatabase.TransactionManager.StartOpenCloseTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(Helper.oDatabase), OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                        Helper.oEditor.WriteMessage("\nchecking: " + id.ObjectClass.Name);

                        Extents3d exts;

                        if (ent.Bounds.HasValue)
                            exts = ent.Bounds.Value;
                        else
                            continue;

                        Point3d bmin = exts.MinPoint;
                        Point3d bmax = exts.MaxPoint;

                        if (BoxesIntersect(themin, themax, bmin, bmax))
                        {
                            Helper.oEditor.WriteMessage("\nintersection with: " + id.ObjectClass.Name);
                            forSelection.Add(id);
                        }
                    }
                    ObjectId[] forSelArr = forSelection.ToArray();
                    Helper.oEditor.SetImpliedSelection(forSelArr);
                    tr.Commit();
                }

                Helper.oEditor.Command("_PLANTISOLATE");

                Application.SetSystemVariable("SNAPMODE", snapmodesaved);
                Application.SetSystemVariable("OSMODE", osmodesaved);
                Application.SetSystemVariable("3DOSMODE", os3dmodesaved);

            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage(e.Message);

            }
            finally { Helper.Terminate(); }
        }


        public static void EndClip()
        {
            Helper.oEditor.Command("_XCLIP", "_ALL", "", "_D");
            Helper.oEditor.Command("_Plantshowall");

        }

        [CommandMethod("prepareBox", CommandFlags.UsePickSet)]
        public static void prepareBox()
        {
            try
            {
                Helper.Initialize();
                verifyLayerIsPresent();

                PromptSelectionResult selUser = Helper.oEditor.GetSelection();

                ObjectId[] objIdArrayUser = new ObjectId[0];

                if (selUser.Status == PromptStatus.OK)
                    objIdArrayUser = selUser.Value.GetObjectIds();

                TypedValue[] filterlist = new TypedValue[2];
                filterlist[0] = new TypedValue(0, "3DSOLID");
                filterlist[1] = new TypedValue(8, "clipbox");
                SelectionFilter filter = new SelectionFilter(filterlist);
                PromptSelectionResult selRes = Helper.oEditor.SelectAll(filter);


                if (selRes.Status != PromptStatus.OK)
                {
                    createBox(new Point3d(1,1,1));
                    selRes = Helper.oEditor.SelectAll(filter);
                }

                ObjectId[] objIdArray = selRes.Value.GetObjectIds();

                if (objIdArray.Length > 1)
                {
                    Helper.oEditor.WriteMessage("\nJust one object (box) is allowed to be on the clipbox layer, cancelling the command..");
                    return;
                }

                using (Transaction tr = Helper.oDatabase.TransactionManager.StartTransaction())
                {


                    Solid3d clipbox = tr.GetObject(objIdArray[0], OpenMode.ForWrite) as Solid3d;
                

                    if (objIdArrayUser.Length == 0)
                    {
                        if (ShowHideBox(true))
                        {
                            EndClip();
                        }
                        else
                            ShowHideBox(false);
                    }
                    else if (objIdArrayUser.Length == 1 || objIdArrayUser.Length == 2)
                    {

                        if (ShowHideBox(true))
                        {
                            ShowHideBox(false);
                        }
                        else
                        {
                            Entity entelem = tr.GetObject(objIdArrayUser[0], OpenMode.ForRead) as Entity;
                            Entity entelem2 = null;
                            if (objIdArrayUser.Length == 2)
                            {
                                entelem2 = tr.GetObject(objIdArrayUser[1], OpenMode.ForRead) as Entity;
                            }
                            Point3d elemsmid = midpoint(entelem, entelem2);
                            Vector3d tomid = elemsmid.GetAsVector();
                            Point3d s = getScales(entelem, entelem2);
                            clipbox.Erase();
                            createBox(s);
                            selRes = Helper.oEditor.SelectAll(filter);
                            objIdArray = selRes.Value.GetObjectIds();
                            clipbox = tr.GetObject(objIdArray[0], OpenMode.ForWrite) as Solid3d;                        
                            clipbox.TransformBy(Matrix3d.Displacement(tomid));                      
                        }

                    }
                    else
                    {
                        Helper.oEditor.WriteMessage("\nnot possible to select more than 2 elements for this command..\n");
                    }



                    tr.Commit();
                }
            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage(e.Message + ", " + e.StackTrace);

            }
            finally { Helper.Terminate(); }
        }

        public static Point3d midpoint(Entity elem, Entity elem2)
        {
            Extents3d elembounds = elem.Bounds.Value;
            if (elem2 != null) elembounds.AddExtents(elem2.Bounds.Value);
            Point3d ebmin = elembounds.MinPoint;
            Point3d ebmax = elembounds.MaxPoint;
            Point3d ebmid = (new LineSegment3d(ebmin, ebmax)).MidPoint;
            return ebmid;
        }

        public static double lenDiagonal(Entity elem)
        {
            Extents3d elembounds = elem.Bounds.Value;
            Point3d ebmin = elembounds.MinPoint;
            Point3d ebmax = elembounds.MaxPoint;
            double lenDiag = ebmin.DistanceTo(ebmax);
            return lenDiag;
        }

        public static Point3d getScales(Entity elem, Entity elem2)
        {
            Extents3d elembounds = elem.Bounds.Value;
            if (elem2 != null) elembounds.AddExtents(elem2.Bounds.Value);
            Point3d ebmin = elembounds.MinPoint;
            Point3d ebmax = elembounds.MaxPoint;
            double sx = Math.Abs(ebmax.X - ebmin.X);
            double sy = Math.Abs(ebmax.Y - ebmin.Y);
            double sz = Math.Abs(ebmax.Z - ebmin.Z);
            if (sx == 0) sx = 1;
            if (sy == 0) sy = 1;
            if (sz == 0) sz = 1;
            Point3d s = new Point3d(sx, sy, sz);
            return s;
        }

        public static void createBox(Point3d s)
        {
            string sLayerName = "clipbox";
            // Get the current document and database
            Document acDoc = Helper.ActiveDocument;
            Database acCurDb = Helper.oDatabase;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;


                // Create a box object
                using (Solid3d thesolid = new Solid3d())
                {
                    thesolid.RecordHistory = true;
                    thesolid.CreateBox(s.X, s.Y, s.Z);
                    thesolid.Layer = sLayerName;

                    acBlkTblRec.AppendEntity(thesolid);
                    acTrans.AddNewlyCreatedDBObject(thesolid, true);
                }

                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }

        }


        public static bool ShowHideBox(bool isHiddenRequest)
        {

            bool isHiddenReply = false;
            // Get the current document and database
            Document acDoc = Helper.ActiveDocument;
            Database acCurDb = Helper.oDatabase;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                                OpenMode.ForRead) as LayerTable;

                string sLayerName = "clipbox";



                LayerTableRecord acLyrTblRec = acTrans.GetObject(acLyrTbl[sLayerName],
                                                OpenMode.ForWrite) as LayerTableRecord;

                // Turn the layer off
                if (!acLyrTblRec.IsOff)
                {
                    isHiddenReply = false;
                    if (!isHiddenRequest)
                        acLyrTblRec.IsOff = true;
                }
                else
                {
                    isHiddenReply = true;
                    if (!isHiddenRequest)
                        acLyrTblRec.IsOff = false;
                }


                acTrans.Commit();
            }

            return isHiddenReply;
        }


        public static void verifyLayerIsPresent()
        {
            string sLayerName = "clipbox";
            // Get the current document and database
            Document acDoc = Helper.ActiveDocument;
            Database acCurDb = Helper.oDatabase;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                                OpenMode.ForRead) as LayerTable;

                if (acLyrTbl.Has(sLayerName) == false)
                {
                    using (LayerTableRecord acLyrTblRec = new LayerTableRecord())
                    {
                        // Assign the layer a name
                        acLyrTblRec.Name = sLayerName;

                        // Upgrade the Layer table for write
                        acLyrTbl.UpgradeOpen();

                        // Append the new layer to the Layer table and the transaction
                        acLyrTbl.Add(acLyrTblRec);
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);

                        // Turn the layer on
                        acLyrTblRec.IsOff = false;

                        acLyrTblRec.Transparency = new Transparency(70);

                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 7);//black/white

                        acTrans.Commit();
                    }
                }

            }

        }

        public static bool BoxesIntersect(Point3d amin, Point3d amax, Point3d bmin, Point3d bmax)
        {
            if (amax.X < bmin.X) return false; // a is left of b
            if (amin.X > bmax.X) return false; // a is right of b
            if (amax.Y < bmin.Y) return false; // a is above b
            if (amin.Y > bmax.Y) return false; // a is below b
            if (amax.Z < bmin.Z) return false; // a is above b
            if (amin.Z > bmax.Z) return false; // a is below b
            return true; // boxes overlap
        }

        public static bool drawingHasXrefs()
        {
            bool hasxrefs = false;

            try
            {
                using (Transaction trans = Helper.oDatabase.TransactionManager.StartTransaction())
                {


                    BlockTable bt = (BlockTable)trans.GetObject(Helper.oDatabase.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in bt)
                    {

                        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(btrId, OpenMode.ForRead);

                        if (btr.IsFromExternalReference)
                        {

                            hasxrefs = true;
                            break;
                        }

                    }

                    trans.Commit();

                }
            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage("Error: Exception while checking for intersections\n");
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                Helper.oEditor.WriteMessage(trace.ToString());
                Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                Helper.oEditor.WriteMessage("message: " + e.Message);
            }

            return hasxrefs;
        }

        [CommandMethod("loadRequiredXrefs")]
        public static void loadRequiredXrefs()
        {
            try
            {
                Helper.Initialize();
                String dwgDirectory = "";
                PromptResult pr = Helper.oEditor.GetString("\nplease enter project dwg path: ");

                if (pr.Status == PromptStatus.OK && Directory.Exists(pr.StringResult))
                {
                    dwgDirectory = pr.StringResult;
                }
                else
                {
                    Helper.oEditor.WriteMessage("\nno dwg path given, using Plant 3D models folder");
                }

                detachallxrefs();

                verifyLayerIsPresent();

                if (ShowHideBox(true))
                {
                    Helper.oEditor.WriteMessage("\nPlease prepare box first and close all project drawings except this one!");
                    return;
                }

                if (drawingHasXrefs())
                {
                    Helper.oEditor.WriteMessage("\nPlease detach all xrefs from this drawing!");
                    return;
                }

                List<string> xrefs = getIntersectingXrefList(dwgDirectory);


                foreach (string xrefPath in xrefs)
                {
                    using (Transaction tr = Helper.oDatabase.TransactionManager.StartOpenCloseTransaction())
                    {
                        string blockname = String.Empty;
                        if (xrefPath.IndexOf("/") != -1)
                        {
                            blockname = xrefPath.Substring(xrefPath.LastIndexOf("/"));
                        }
                        else
                        {
                            blockname = xrefPath.Substring(xrefPath.LastIndexOf("\\"));
                        }
                        ObjectId xrefObj = Helper.oDatabase.OverlayXref(xrefPath, blockname);
                        if (!xrefObj.IsNull)
                        {
                            using (BlockReference br = new BlockReference(new Point3d(0, 0, 0), xrefObj))
                            {
                                BlockTableRecord btr = tr.GetObject(Helper.oDatabase.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                btr.AppendEntity(br);
                                tr.AddNewlyCreatedDBObject(br, true);
                            }

                        }
                        tr.Commit();
                    }
                }




            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage("Error: Exception while checking for intersections\n");
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                Helper.oEditor.WriteMessage(trace.ToString());
                Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                Helper.oEditor.WriteMessage("message: " + e.Message);

            }
            finally
            {
                Helper.Terminate();
            }

        }


        public static List<string> getIntersectingXrefList(string dwgdir)
        {
            List<string> resultList = new List<string>();
            try
            {
                Extents3d thext = getBoxExtends();
                Point3d amin = thext.MinPoint;
                Point3d amax = thext.MaxPoint;
                string[] dwgList = null;

                if (dwgdir.Equals(String.Empty))
                {
                    System.Collections.Generic.List<PnPProjectDrawing> dwgListp = Helper.PlantProject.GetPnPDrawingFiles();
                    List<string> dwgListp2 = new List<string>();
                    foreach (PnPProjectDrawing dwgp in dwgListp)
                    {
                        dwgListp2.Add(dwgp.ResolvedFilePath);
                    }
                    dwgList = dwgListp2.ToArray();
                }
                else
                {
                    dwgList = Directory.GetFiles(dwgdir, "*.dwg");
                }

                //foreach (PnPProjectDrawing dwg in dwgList)
                foreach (string thepath in dwgList)
                {
                    if (thepath.Equals(Helper.ActiveDocument.Name))
                        continue;

                    double minx = Double.MaxValue;
                    double miny = Double.MaxValue;
                    double minz = Double.MaxValue;
                    double maxx = Double.MinValue;
                    double maxy = Double.MinValue;
                    double maxz = Double.MinValue;
                    bool writebounds = true;

                    Helper.oEditor.WriteMessage("\nchecking file: " + thepath);
                    int numobj = 0;

                    if (File.Exists(thepath + ".bounds"))
                    {
                        DateTime dwgtime = File.GetLastWriteTime(thepath);
                        DateTime cachedtime = File.GetLastWriteTime(thepath + ".bounds");

                        if (dwgtime < cachedtime)
                        {

                            string[] blines = File.ReadAllLines(thepath + ".bounds");
                            if (blines.Length == 6)
                            {
                                if (!BoxesIntersect(amin, amax, new Point3d(Convert.ToDouble(blines[0]), Convert.ToDouble(blines[1]), Convert.ToDouble(blines[2])), new Point3d(Convert.ToDouble(blines[3]), Convert.ToDouble(blines[4]), Convert.ToDouble(blines[5]))))
                                {
                                    continue;
                                }

                            }

                        }
                    }

                    using (Database sideDb = new Database(false, true))
                    {


                        // read the dwg file
                        //file open cannot be read
                        //glue attachement loading slows down the tool (model download)
                        try
                        {
                            sideDb.ReadDwgFile(thepath, FileOpenMode.OpenForReadAndAllShare, true, null);
                        }
                        catch (System.Exception e)
                        {
                            Helper.oEditor.WriteMessage("Error: Exception while reading file: " + thepath + " continue with next..\n");
                            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                            Helper.oEditor.WriteMessage(trace.ToString());
                            Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                            Helper.oEditor.WriteMessage("message: " + e.Message);
                            continue;
                        }

                        // close the file and make sure we have read everything in

                        sideDb.CloseInput(true);
                        using (Transaction sidetr = sideDb.TransactionManager.StartOpenCloseTransaction())
                        {
                            BlockTableRecord btr = (BlockTableRecord)sidetr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(sideDb), OpenMode.ForRead);


                            foreach (ObjectId id in btr)
                            {
                                if (id.ObjectClass.Name.Equals("AcDbBlockReference"))
                                {
                                    BlockReference blockref = (BlockReference)sidetr.GetObject(id, OpenMode.ForRead);
                                    BlockTableRecord blockbtr = (BlockTableRecord)sidetr.GetObject(blockref.BlockTableRecord, OpenMode.ForRead);
                                    if (blockbtr.IsFromExternalReference) continue;
                                }

                                if (id.ObjectClass.Name.Equals("AcPpDb3dConnector"))
                                    continue;

                                if (id.ObjectClass.Name.Equals("AcDbCircle"))
                                    continue;
                                //AcDbZombieEntity (proxy)

                                ++numobj;

                                Entity ent = (Entity)sidetr.GetObject(id, OpenMode.ForRead);
                                //Helper.oEditor.WriteMessage("\nchecking: " + id.ObjectClass.Name);

                                Extents3d exts;

                                if (ent.Bounds.HasValue)
                                    exts = ent.Bounds.Value;
                                else
                                    continue;

                                Point3d bmin = exts.MinPoint;
                                Point3d bmax = exts.MaxPoint;

                                if (BoxesIntersect(amin, amax, bmin, bmax))
                                {
                                    Helper.oEditor.WriteMessage("\nintersection with: " + id.ObjectClass.Name);
                                    resultList.Add(thepath);
                                    writebounds = false;
                                    break;
                                }

                                if (bmin.X < minx) minx = bmin.X;
                                if (bmin.Y < miny) miny = bmin.Y;
                                if (bmin.Z < minz) minz = bmin.Z;
                                if (bmax.X > maxx) maxx = bmax.X;
                                if (bmax.Y > maxy) maxy = bmax.Y;
                                if (bmax.Z > maxz) maxz = bmax.Z;

                                if (bmax.X < minx) minx = bmax.X;
                                if (bmax.Y < miny) miny = bmax.Y;
                                if (bmax.Z < minz) minz = bmax.Z;
                                if (bmin.X > maxx) maxx = bmin.X;
                                if (bmin.Y > maxy) maxy = bmin.Y;
                                if (bmin.Z > maxz) maxz = bmin.Z;
                            }
                            //Helper.oEditor.WriteMessage("sel: " + xyselres.Status);
                            sidetr.Commit();
                        }
                    }
                    //write boundsfile
                    if (writebounds && numobj > 0)
                    {
                        File.WriteAllText(thepath + ".bounds", minx + "\n" + miny + "\n" + minz + "\n" + maxx + "\n" + maxy + "\n" + maxz);
                    }
                }
            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage("Error: Exception while checking for intersections\n");
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                Helper.oEditor.WriteMessage(trace.ToString());
                Helper.oEditor.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                Helper.oEditor.WriteMessage("message: " + e.Message);

            }
            finally { }
            return resultList;
        }


    }
}
