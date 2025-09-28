namespace openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;

/// <summary>
/// Context object that holds the migration state during comment processing.
/// </summary>
public class MigrationContext
{
    public List<ThreadedCommentWithContext> SortedComments { get; set; } = [];
    public HashSet<string> ProcessedCells { get; } = [];
}
