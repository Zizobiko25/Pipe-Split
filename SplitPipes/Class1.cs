using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace SplitPipes
{
  [Autodesk.Revit.Attributes.Transaction( Autodesk.Revit.Attributes.TransactionMode.Manual )]
  public class SplitPipes : IExternalCommand
  {
    // The main Execute method (inherited from IExternalCommand) must be public
    public Result Execute( ExternalCommandData revit,
        ref string message, ElementSet elements )
    {
      try
      {
        // Verify if the active document is null
        UIDocument activeDoc = revit.Application.ActiveUIDocument;
        if( activeDoc == null )
        {
          TaskDialog.Show( "No Active Document", "There's no active document in Revit.", TaskDialogCommonButtons.Ok );
          return Autodesk.Revit.UI.Result.Failed;
        }

        // Verify the number of selected elements
        ElementSet selElements = new ElementSet();
        foreach( ElementId elementId in activeDoc.Selection.GetElementIds() )
        {
          selElements.Insert( activeDoc.Document.GetElement( elementId ) );
        }
        if( selElements.Size != 1 )
        {
          message = "Please select ONLY one element from current project.";
          return Autodesk.Revit.UI.Result.Failed;
        }

        // Get the selected element
        Element selectedElement = null;
        foreach( Element element in selElements )
        {
          selectedElement = element;
          break;
        }


        MEPSystem system = ExtractMechanicalOrPipingSystem( selectedElement );
        if( system == null )
        {
          message = "Error! Check the system conectivity";
          return Autodesk.Revit.UI.Result.Failed;
        }

        List<Element> pipes = new List<Element>();
        FilteredElementCollector collector = new FilteredElementCollector( activeDoc.Document );
        collector.OfClass( typeof( PipeType ) );
        PipeType pipeType = collector.FirstElement() as PipeType;

        Selection sel = activeDoc.Selection;

        foreach( ElementId eId in sel.GetElementIds() )
        {
          Element e = activeDoc.Document.GetElement( eId );
          Pipe c = e as Pipe;
          pipeType = c.PipeType;

          if( null != c )
          {
            pipes.Add( c );
            if( 2 == pipes.Count )
            {
              break;
            }
          }
        }

        using( Transaction tx = new Transaction( activeDoc.Document ) )
        {
          tx.Start( "split pipe" );
          ElementId systemtype = system.GetTypeId();
          SplitPipe( pipes[ 0 ], system, activeDoc, systemtype, pipeType );
          tx.Commit();
        }
        return Autodesk.Revit.UI.Result.Succeeded;
      }
      catch( Exception ex )
      {
        message = ex.Message;
        return Autodesk.Revit.UI.Result.Failed;
      }
    }

    private void SplitPipe( Element segment, MEPSystem _system, UIDocument _activeDoc, ElementId _systemtype, PipeType _pipeType )
    {
      ElementId levelId = segment.get_Parameter( BuiltInParameter.RBS_START_LEVEL_PARAM ).AsElementId();
      // system.LevelId;
      ElementId systemtype = _system.GetTypeId();

      Curve c1 = (segment.Location as LocationCurve).Curve;// selecting one pipe and taking its location.

      //Pipe diameter
      double pipeDia = UnitUtils.ConvertFromInternalUnits( segment.get_Parameter( BuiltInParameter.RBS_PIPE_DIAMETER_PARAM ).AsDouble(), DisplayUnitType.DUT_MILLIMETERS );

      //Standard length
      double l = 6000;

      //Coupling length
      double fittinglength = (1.1 * pipeDia + 14.4);

      // finding the length of the selected pipe.
      double len = UnitUtils.ConvertFromInternalUnits( segment.get_Parameter( BuiltInParameter.CURVE_ELEM_LENGTH ).AsDouble(), DisplayUnitType.DUT_MILLIMETERS );

      if( len <= l )
        return;

      var startPoint = c1.GetEndPoint( 0 );
      var endPoint = c1.GetEndPoint( 1 );

      XYZ splitpoint = (endPoint - startPoint) * (l / len);

      var newpoint = startPoint + splitpoint;

      Pipe pp = segment as Pipe;

      // Find two connectors which pipe's two ends connector connected to. 
      Connector startConn = FindConnectedTo( pp, startPoint );
      Connector endConn = FindConnectedTo( pp, endPoint );

      // creating first pipe 
      Pipe pipe = null;
      if( null != _pipeType )
      {
        pipe = Pipe.Create( _activeDoc.Document, _pipeType.Id, levelId, startConn, newpoint );
      }

      Connector conn1 = FindConnector( pipe, newpoint );

      //Check + fitting
      XYZ fittingend = (endPoint - startPoint) * ((l + (fittinglength / 2)) / len);

      //New point after the fitting gap
      var endOfFitting = startPoint + fittingend;

      Pipe pipe1 = Pipe.Create( _activeDoc.Document, systemtype, _pipeType.Id, levelId, endOfFitting, endPoint );

      // Copy parameters from previous pipe to the following Pipe. 
      CopyParameters( pipe, pipe1 );
      Connector conn2 = FindConnector( pipe1, endOfFitting );
      _ = _activeDoc.Document.Create.NewUnionFitting( conn1, conn2 );

      if( null != endConn )
      {
        Connector pipeEndConn = FindConnector( pipe1, endPoint );
        pipeEndConn.ConnectTo( endConn );

      }

      ICollection<Autodesk.Revit.DB.ElementId> deletedIdSet = _activeDoc.Document.Delete( segment.Id );
      if( 0 == deletedIdSet.Count )
      {
        throw new Exception( "Deleting the selected elements in Revit failed." );
      }


      if( UnitUtils.ConvertFromInternalUnits( pipe1.get_Parameter( BuiltInParameter.CURVE_ELEM_LENGTH ).AsDouble(), DisplayUnitType.DUT_MILLIMETERS ) > l )
      {
        SplitPipe( pipe1, _system, _activeDoc, _systemtype, _pipeType );
      }
    }

    private static Connector FindConnector( Pipe pipe, Autodesk.Revit.DB.XYZ conXYZ )
    {
      ConnectorSet conns = pipe.ConnectorManager.Connectors;
      foreach( Connector conn in conns )
      {
        if( conn.Origin.IsAlmostEqualTo( conXYZ ) )
        {
          return conn;
        }
      }
      return null;
    }

    private static Connector FindConnectedTo( Pipe pipe, Autodesk.Revit.DB.XYZ conXYZ )
    {
      Connector connItself = FindConnector( pipe, conXYZ );
      ConnectorSet connSet = connItself.AllRefs;
      foreach( Connector conn in connSet )
      {
        if( conn.Owner.Id.IntegerValue != pipe.Id.IntegerValue &&
            conn.ConnectorType == ConnectorType.End )
        {
          return conn;
        }
      }
      return null;
    }

    private static void CopyParameters( Pipe source, Pipe target )
    {
      double diameter = source.get_Parameter( BuiltInParameter.RBS_PIPE_DIAMETER_PARAM ).AsDouble();
      target.get_Parameter( BuiltInParameter.RBS_PIPE_DIAMETER_PARAM ).Set( diameter );
    }

    /// <summary>
    /// Get the mechanical or piping system from selected element
    /// </summary>
    /// <param name="selectedElement">Selected element</param>
    /// <returns>The extracted mechanical or piping system. Null if no expected system is found.</returns>
    private MEPSystem ExtractMechanicalOrPipingSystem( Element selectedElement )
    {
      MEPSystem system = null;

      if( selectedElement is MEPSystem )
      {
        if( selectedElement is MechanicalSystem || selectedElement is PipingSystem )
        {
          system = selectedElement as MEPSystem;
          return system;
        }
      }
      else // Selected element is not a system
      {
        // If selected element is a family instance, iterate its connectors and get the expected system
        if( selectedElement is FamilyInstance fi )
        {
          MEPModel mepModel = fi.MEPModel;
          ConnectorSet connectors = null;
          try
          {
            connectors = mepModel.ConnectorManager.Connectors;
          }
          catch( System.Exception )
          {

          }

          system = ExtractSystemFromConnectors( connectors );
        }
        else
        {
          //
          // If selected element is a MEPCurve (e.g. pipe or duct), 
          // iterate its connectors and get the expected system

          if( selectedElement is MEPCurve mepCurve )
          {
            ConnectorSet connectors = mepCurve.ConnectorManager.Connectors;
            system = ExtractSystemFromConnectors( connectors );
          }
        }
      }

      return system;
    }

    /// <summary>
    /// Get the mechanical or piping system from the connectors of selected element
    /// </summary>
    /// <param name="connectors">Connectors of selected element</param>
    /// <returns>The found mechanical or piping system</returns>
    static private MEPSystem ExtractSystemFromConnectors( ConnectorSet connectors )
    {
      MEPSystem system = null;

      if( connectors == null || connectors.Size == 0 )
      {
        return null;
      }

      // Get well-connected mechanical or piping systems from each connector
      List<MEPSystem> systems = new List<MEPSystem>();
      foreach( Connector connector in connectors )
      {
        MEPSystem tmpSystem = connector.MEPSystem;
        if( tmpSystem == null )
        {
          continue;
        }

        if( tmpSystem is MechanicalSystem ms )
        {
          if( ms.IsWellConnected )
          {
            systems.Add( tmpSystem );
          }
        }
        else
        {
          if( tmpSystem is PipingSystem ps && ps.IsWellConnected )
          {
            systems.Add( tmpSystem );
          }
        }
      }

      // If more than one system is found, get the system contains the most elements
      int countOfSystem = systems.Count;
      if( countOfSystem != 0 )
      {
        int countOfElements = 0;
        foreach( MEPSystem sys in systems )
        {
          if( sys.Elements.Size > countOfElements )
          {
            system = sys;
            countOfElements = sys.Elements.Size;
          }
        }
      }

      return system;
    }
  }
}
