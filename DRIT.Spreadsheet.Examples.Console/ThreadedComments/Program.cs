using System.Linq;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
var author = workbook.ThreadedCommentAuthors.Add("Alice");
var parent = worksheet.ThreadedComments.Add("A1", "Please review this value.", author);
worksheet.ThreadedComments.AddReply(parent, "Reviewed and approved.", author);
var secondAuthor = workbook.ThreadedCommentAuthors.Add("Bob");
worksheet.ThreadedComments.Add("B2", "This is a follow-up.", secondAuthor).Resolve();

var path = ExampleSupport.OutputPath("ThreadedComments.xlsx");
workbook.SaveAs(path);
var loaded = Workbook.Load(path);
var loadedWorksheet = loaded.Worksheets[0];
var comments = loadedWorksheet.ThreadedComments.GetByCell("A1").ToList();
ExampleSupport.Require(loaded.ThreadedCommentAuthors.Count == 2, "Threaded comment authors were not preserved.");
ExampleSupport.Require(comments.Count == 2, "The threaded comment reply was not preserved.");
ExampleSupport.Require(comments.Any(comment => comment.Parent == null && comment.Text == "Please review this value."), "The top-level threaded comment was not preserved.");
ExampleSupport.Require(comments.Any(comment => comment.Parent != null && comment.Text == "Reviewed and approved."), "The threaded comment reply was not preserved.");
ExampleSupport.Require(loadedWorksheet.ThreadedComments.GetByCell("B2").Single().Done, "The resolved threaded comment was not preserved.");
Console.WriteLine("Threaded comments, replies, authors, and resolved state preserved.");
