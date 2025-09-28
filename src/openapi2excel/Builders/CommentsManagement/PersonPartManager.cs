using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Packaging;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Factory for managing person parts in OpenXML workbooks for threaded comments.
/// </summary>
public class PersonPartManager
{
    /// <summary>
    /// Ensures PersonPart exists for all comment authors using official SDK pattern.
    /// </summary>
    public void EnsurePersonsPartExistsForComments(
        WorkbookPart workbookPart, 
        List<ThreadedCommentWithContext> comments, 
        string existingWorkbookPath)
    {
        var personPart = workbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
        if (personPart == null)
        {
            personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
            personPart.PersonList = new PersonList();
        }

        // Add persons for each unique author
        foreach (var comment in comments)
        {
            if (comment.Comment.PersonId?.Value != null)
            {
                var personId = comment.Comment.PersonId.Value;
                
                // Check if person already exists
                var existingPerson = personPart.PersonList.Elements<Person>()
                    .FirstOrDefault(p => p.Id?.Value == personId);
                
                if (existingPerson == null)
                {
                    // Extract person from source or create default
                    var person = ExtractPersonFromSourceWorkbook(existingWorkbookPath, personId);
                    personPart.PersonList.AppendChild(new Person
                    {
                        Id = personId,
                        DisplayName = person.DisplayName ?? "OpenAPI2Excel User",
                        ProviderId = person.ProviderId ?? "Excel",
                        UserId = person.UserId ?? "user@example.com"
                    });
                }
            }
        }
    }

    /// <summary>
    /// Extracts a specific Person by ID from the source workbook using OpenXML objects.
    /// Returns null if no persons part exists or the person is not found.
    /// </summary>
    private static Person ExtractPersonFromSourceWorkbook(string existingWorkbookPath, string personId)
    {
        var defaultPerson = new Person
        {
            Id = personId,
            DisplayName = "Comment Author",
            ProviderId = "Excel"
        };

        try
        {
            using var sourceDocument = SpreadsheetDocument.Open(existingWorkbookPath, false);
            var sourceWorkbookPart = sourceDocument.WorkbookPart;
            if (sourceWorkbookPart == null)
            {
                return defaultPerson;
            }

            var sourcePersonPart = sourceWorkbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
            if (sourcePersonPart?.PersonList == null)
            {
                return defaultPerson;
            }

            // Find the specific person by ID
            var person = sourcePersonPart.PersonList.Elements<Person>()
                .FirstOrDefault(p => p.Id?.Value == personId);

            if (person != null)
            {
                return (Person)person.Clone();
            }
            else
            {
                return defaultPerson;
            }
        }
        catch (Exception)
        {
            return defaultPerson;
        }
    }
}