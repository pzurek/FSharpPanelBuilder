namespace FSharpPanelBuilder

open System
open System.Collections.Generic

open Autodesk.Revit
open Autodesk.Revit.UI
open Autodesk.Revit.Attributes
open Autodesk.Revit.DB
open Autodesk.Revit.UI.Selection

type WallSelectionFilter() =
  class
    interface ISelectionFilter with
      member x.AllowElement(element) =
        match element with
        | :? Wall as element -> true
        | _ -> false

      member x.AllowReference(reference, xyz) =
        false
  end

[<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>]
type Command() =
  class
    interface IExternalCommand with
      member this.Execute(commandData, message : string byref, elements) =
        try
          use uiApp = commandData.Application
          use app  = uiApp.Application
          use uiDoc = uiApp.ActiveUIDocument
          use doc  = uiDoc.Document
          
          let reference = uiDoc.Selection.PickObject(ObjectType.Element, new WallSelectionFilter(), "Select a wall to split into panels")

          let wall = doc.GetElement(reference.ElementId) :?> Wall
          let locationCurve = wall.Location :?> LocationCurve 
          let line = locationCurve.Curve :?> Line

          use transaction = new Transaction(doc)
          let status = transaction.Start("Building panels")
          let wallList = new List<ElementId>(1)
          wallList.Add(reference.ElementId)

          PartUtils.CreateParts(doc, wallList)
          doc.Regenerate()
          let parts = PartUtils.GetAssociatedParts(doc, wall.Id, false, false)
          let divisions = 15
          let origin = line.Origin
          let delta = line.Direction.Multiply(line.Length / (float)divisions)
          let shiftDelta = Transform.get_Translation(delta)
          let rotation = Transform.get_Rotation(origin, XYZ.BasisZ, 0.5 * Math.PI)
          let wallWidthVector = rotation.OfVector(line.Direction.Multiply(2. * wall.Width))
          let mutable intersectionLine = Line.CreateBound(origin + wallWidthVector, origin - wallWidthVector)
          let curveArray = new List<Curve>()

          for i = 1 to divisions do
            intersectionLine <- intersectionLine.get_Transformed(shiftDelta) :?> Line
            curveArray.Add(intersectionLine)

          let divisionSketchPlane = SketchPlane.Create(doc, new Plane(XYZ.BasisZ, line.Origin))
          let intersectionElementsIds = new List<ElementId>()
          let partMaker = PartUtils.DivideParts(doc, parts, intersectionElementsIds, curveArray, divisionSketchPlane.Id)
          doc.ActiveView.PartsVisibility <- PartsVisibility.ShowPartsOnly
          transaction.Commit() |> ignore

          Autodesk.Revit.UI.Result.Succeeded

        with ex ->
          TaskDialog.Show("F# Panel Builder", ex.Message) |> ignore
          Autodesk.Revit.UI.Result.Failed

  end