using DocumentFormat.OpenXml.Packaging;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Factory for creating VML Drawing parts using the official Microsoft SDK patterns.
/// </summary>
public class VmlDrawingFactory
{
    private readonly CommentTargetResolver _targetResolver;

    public VmlDrawingFactory(CommentTargetResolver targetResolver)
    {
        _targetResolver = targetResolver ?? throw new ArgumentNullException(nameof(targetResolver));
    }

    /// <summary>
    /// Creates VML Drawing Part using the exact official SDK pattern.
    /// This is the proven working VML that Excel accepts.
    /// </summary>
    public void CreateVmlDrawingPartUsingOfficialPattern(
        WorksheetPart worksheetPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Check if ClosedXML already created a VML part
        var existingVmlPart = worksheetPart.GetPartsOfType<VmlDrawingPart>().FirstOrDefault();
        VmlDrawingPart vmlDrawingPart;
        
        if (existingVmlPart != null)
        {
            // Use the existing VML part but replace its content
            vmlDrawingPart = existingVmlPart;
        }
        else
        {
            // Create new VML part with specific relationship ID like official example
            vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>("rId1");
        }

        using var writer = new System.Xml.XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8);
        // Use the EXACT VML from the official SDK example that works
        string vmlContent = @"<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">
                <o:shapelayout v:ext=""edit"">
                    <o:idmap v:ext=""edit"" data=""1""/>
                </o:shapelayout>
                <v:shapetype id=""_x0000_t202"" coordsize=""21600,21600"" o:spt=""202"" path=""m,l,21600r21600,l21600,xe"">
                    <v:stroke joinstyle=""miter""/>
                    <v:path gradientshapeok=""t"" o:connecttype=""rect""/>
                </v:shapetype>";

        int shapeId = 1025; // Use official example's starting shape ID

        // Create VML shape for each root comment
        foreach (var comment in comments.Where(c => c.IsRootComment))
        {
            if (!TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out string targetCellReference))
                continue;

            // Extract row and column for VML anchor (0-based for VML)
            var row = _targetResolver.ExtractRowFromCellReference(targetCellReference) - 1;
            var col = _targetResolver.ExtractColumnIndexFromCellReference(targetCellReference);

            // Use EXACT VML shape pattern from official example - CRITICAL: no space after semicolon
            vmlContent += $@"
                <v:shape id=""_x0000_s{shapeId}"" type=""#_x0000_t202"" style=""position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden"" fillcolor=""#ffffe1"" o:insetmode=""auto"">
                    <v:fill color2=""#ffffe1""/>
                    <v:shadow on=""t"" color=""black"" obscured=""t""/>
                    <v:path o:connecttype=""none""/>
                    <v:textbox style=""mso-direction-alt:auto"">
                        <div style=""text-align:left""></div>
                    </v:textbox>
                    <x:ClientData ObjectType=""Note"">
                        <x:MoveWithCells/>
                        <x:SizeWithCells/>
                        <x:Anchor>1, 15, {row}, 2, 3, 15, {row + 3}, 16</x:Anchor>
                        <x:AutoFill>False</x:AutoFill>
                        <x:Row>{row}</x:Row>
                        <x:Column>{col}</x:Column>
                    </x:ClientData>
                </v:shape>";

            shapeId++;
        }

        vmlContent += "</xml>";
        writer.WriteRaw(vmlContent);
        writer.Flush();
    }

    private bool TryGetTargetCellForThreadedComment(
        ThreadedCommentWithContext comment, 
        List<WorksheetOpenApiMapping> newWorkbookMappings, 
        out string targetCellReference)
    {
        return _targetResolver.TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out targetCellReference);
    }
}