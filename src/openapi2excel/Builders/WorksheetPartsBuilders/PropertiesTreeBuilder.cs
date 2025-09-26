using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.CustomXml;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class PropertiesTreeBuilder(
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
{
   private readonly int _attributesColumnIndex = attributesColumnIndex + 2;
   protected OpenApiDocumentationOptions Options { get; } = options;
   protected IXLWorksheet Worksheet { get; } = worksheet;
   private RowPointer ActualRow { get; set; } = null!;
   protected XLColor HeaderBackgroundColor => XLColor.LightGray;

   public void AddPropertiesTreeForMediaTypes(RowPointer actualRow, IDictionary<string, OpenApiMediaType> mediaTypes, OpenApiDocumentationOptions options, Anchor mappingAnchor)
   {
      ActualRow = actualRow;
      foreach (var mediaType in mediaTypes)
      {
         var bodyFormatRowPointer = ActualRow.Copy();
         Worksheet.Cell(ActualRow, 1).SetTextBold($"{WorksheetLanguage.Request.BodyFormat}: {mediaType.Key}")
            .MapRow(AnchorGenerator.GenerateBodyFormatAnchor(mappingAnchor));
         ActualRow.MoveNext();

         if (mediaType.Value.Schema != null)
         {
            using (var _ = new Section(Worksheet, ActualRow))
            {
               var columnCount = AddPropertiesTree(ActualRow, mediaType.Value.Schema, options, mappingAnchor);
               Worksheet.Cell(bodyFormatRowPointer, 1).SetBackground(columnCount, HeaderBackgroundColor);
               ActualRow.MovePrev();
            }
            ActualRow.MoveNext(2);
         }
         else
         {
            ActualRow.MoveNext();
         }
      }
   }

   public int AddPropertiesTree(RowPointer actualRow, OpenApiSchema schema, OpenApiDocumentationOptions options, Anchor mappingAnchor)
   {
      ActualRow = actualRow;
      var columnCount = AddSchemaDescriptionHeader(mappingAnchor);
      var startLevel = CorrectRootElementIfArray(schema, mappingAnchor) ? 2 : 1;
      AddProperties(schema, startLevel, options, mappingAnchor);
      return columnCount;
   }

   protected bool CorrectRootElementIfArray(OpenApiSchema schema, Anchor mappingAnchor)
   {
      if (schema.Items == null)
         return false;

      AddPropertyRow("<array>", schema, false, 1, mappingAnchor);
      return true;
   }

   protected void AddProperties(OpenApiSchema? schema, int level, OpenApiDocumentationOptions options, Anchor mappingAnchor)
   {
      if (schema == null)
         return;

      if (schema.Items != null)
      {
         AddPropertiesForArray(schema, level, options, mappingAnchor);
      }

      if (schema.AllOf.Any())
      {
         // Add all subschemas on the current level
         schema.AllOf.ForEach(subschema =>
         {
            AddProperties(subschema, level, options, mappingAnchor);
         });
      }

      if (schema.AnyOf.Count == 1)
      {
         // If there is only one subschema, we can treat it as a single schema
         AddProperties(schema.AnyOf[0], level, options, mappingAnchor);
      }
      else if (schema.AnyOf.Any())
      {
         // Otherwise, add each subschema one level below, indicating their disjunction composition
         schema.AnyOf.ForEach(subschema =>
         {
            AddPropertyRow("<anyOf>", subschema, false, level, mappingAnchor);
            AddProperties(subschema, level + 1, options, mappingAnchor);
         });
      }

      if (schema.OneOf.Count == 1)
      {
         // If there is only one subschema, we can treat it as a single schema
         AddProperties(schema.OneOf[0], level, options, mappingAnchor);
      }
      else if (schema.OneOf.Any())
      {
         // Otherwise, add each subschema one level below, indicating their disjunction composition
         schema.OneOf.ForEach(subschema =>
         {
            AddPropertyRow("<oneOf>", subschema, false, level, mappingAnchor);
            AddProperties(subschema, level + 1, options, mappingAnchor);
         });
      }
      foreach (var property in schema.Properties)
      {
         AddProperty(property.Key, property.Value, schema.Required.Contains(property.Key), level, options, mappingAnchor);
      }
   }

   private void AddPropertiesForArray(OpenApiSchema schema, int level, OpenApiDocumentationOptions options, Anchor mappingAnchor)
   {
      if (schema.Items.Properties.Any() || schema.Items.AllOf.Any() || schema.Items.AnyOf.Any() || schema.Items.OneOf.Any())
      {
         // An array of either object properties or subschemas
         AddProperties(schema.Items, level, options, mappingAnchor);
      }
      else
      {
         // if array contains simple type items
         AddProperty("<value>", schema.Items, false, level, options, mappingAnchor);
      }
   }

   protected void AddProperty(string name, OpenApiSchema? schema, bool required, int level, OpenApiDocumentationOptions options, Anchor mappingAnchor)
   {
      if (schema == null || level >= options.MaxDepth)
      {
         return;
      }

      AddPropertyRow(name, schema, required, level++, mappingAnchor);
      AddProperties(schema, level, options, AnchorGenerator.GeneratePropertyAnchor(mappingAnchor, name));
   }

   private void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, bool required, int propertyLevel, Anchor mappingAnchor)
   {
      const int startColumn = 1;
      Worksheet.Cell(ActualRow, startColumn).SetBackground(propertyLevel - 1, HeaderBackgroundColor);
      var propertyAnchor = AnchorGenerator.GeneratePropertyAnchor(mappingAnchor, propertyName);
      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      schemaDescriptor.AddNameValue(propertyName, ActualRow, propertyLevel, propertyAnchor);
      schemaDescriptor.AddSchemaDescriptionValues(propertySchema, required, ActualRow, _attributesColumnIndex, propertyAnchor);
      ActualRow.MoveNext();
   }

   protected int AddSchemaDescriptionHeader(Anchor mappingAnchor)
   {
      const int startColumn = 1;

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      schemaDescriptor.AddNameHeader(ActualRow, startColumn);
      var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, _attributesColumnIndex, mappingAnchor);

      Worksheet.Cell(ActualRow, startColumn)
         .MapRow(AnchorGenerator.GenerateSchemaDescriptionHeaderAnchor(mappingAnchor))
         .SetBackground(lastUsedColumn, HeaderBackgroundColor)
         .SetBottomBorder(lastUsedColumn);

      ActualRow.MoveNext();
      return lastUsedColumn;
   }
}