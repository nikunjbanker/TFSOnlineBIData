using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Discussion.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace TFSOnlineBIData
{
  class Program
  {
    const string  URL_TO_TFS_COLLECTION = "https://decos.visualstudio.com/DefaultCollection";
    static void Main(string[] args)
    {
      BasicAuthCredential bAuth = new BasicAuthCredential(new NetworkCredential("nikunj", "decos@123"));

      var tfs = new TfsTeamProjectCollection(new Uri(URL_TO_TFS_COLLECTION), new TfsClientCredentials(bAuth));
      tfs.Authenticate();

      var store = tfs.GetService<WorkItemStore>();

      var versionStore = tfs.GetService<Microsoft.TeamFoundation.VersionControl.Client.VersionControlServer>();

      StringBuilder swFileData = new StringBuilder();
      //Handler, WorkItemId, ReviewRequestId, Comment, Author, CommentDate {NewLine}
      swFileData.Append(string.Format("{0}, {1}, {2}, {3}, {4}, {5}{6}", "Handler", "WorkItemId", "ReviewRequestId", "Comment", "Author", "CommentDate", Environment.NewLine));

      StringBuilder swReviewResponseData = new StringBuilder();
      //"Handler", "WorkItemId", "ReviewRequestId", "ReviewResponseId", "ClosedBy", "ClosedStatus", "ClosedDate"
      swReviewResponseData.Append(string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}{7}", "Handler", "WorkItemId", "ReviewRequestId", "ReviewResponseId", "ClosedBy", "ClosedStatus", "ClosedDate", Environment.NewLine));

      StringBuilder swReviewRequestData = new StringBuilder();
      //ID, Work Item Type, Title, Assigned To, State, Effort, Tags, Area Path, Changed Date, Closed Date, Created Date, Iteration Path, Remaining Work, ReviewRequestCount
      string sReviewRequestLine = "ID, Work Item Type, Title, Assigned To, State, Effort, Tags, Area Path, Changed Date, Closed Date, Created Date, Iteration Path, Remaining Work, ReviewRequestCount";
      string sReviewRequestHeaderFormat = string.Empty;
      int iColumnIndex = 0;
      foreach (var column in sReviewRequestLine.Split(','))
      {
        if (iColumnIndex > 0) sReviewRequestHeaderFormat += ", ";
        sReviewRequestHeaderFormat += "{" + iColumnIndex + "}";
        iColumnIndex++;
      }
      sReviewRequestHeaderFormat += "{" + iColumnIndex + "}";
      swReviewRequestData.Append(string.Format(sReviewRequestHeaderFormat, "ID" , "Work Item Type", "Title"
                                                , "Assigned To", "State", "Effort", "Tags", "Area Path", "Changed Date", "Closed Date", "Created Date"
                                                , "Iteration Path", "Remaining Work", "ReviewRequestCount", Environment.NewLine));

      var queryText = "SELECT [System.Id] FROM WorkItems WHERE [System.WorkItemType] In ('Product Backlog Item','Bug') And [System.CreatedDate] >= '4/1/2015'" +
                               "And [System.State] <> 'Removed' Order By [System.Id], [System.ChangedDate]";
      var query = new Query(store, queryText);

      var result = query.RunQuery().OfType<WorkItem>().ToList();

      //To get all sub task for Issue & Internal Support - PBI
      Parallel.ForEach(result, (workitem) =>
      {
        AddTaskToResult(workitem, store, result);
      });

      //Parallel.ForEach(result, (workitem) =>   
      foreach (var workitem in result)
      {
        List<WorkItem> wicReviewRequests = GetReviewRequests(store, workitem);
        string sWorkItemNewLine = GetWorkItemLine(sReviewRequestHeaderFormat, workitem, wicReviewRequests);
        swReviewRequestData.Append(sWorkItemNewLine);
        //Console.WriteLine(sRequestNewLine);

        if (wicReviewRequests != null)
        {
          foreach (var reviewrequest in wicReviewRequests)
          {
            try
            {
              #region code review comments
              //1. Get code review comments from review request
              List<CodeReviewComment> comments = GetCodeReviewComments(reviewrequest.Id);
              if (comments != null && comments.Count > 0)
              {
                //Create new line for each comments (CSV)
                foreach (var comment in comments)
                {
                  try
                  {
                    //Handler, WorkItemId, ReviewRequestId, Comment, Author, CommentDate {NewLine}
                    string sNewLine = string.Format("{0}, {1}, {2}, {3}, {4}, {5}{6}", reviewrequest.CreatedBy, workitem.Id, reviewrequest.Id, comment.Comment, comment.Author, comment.PublishDate, Environment.NewLine);
                    swFileData.Append(sNewLine);
                    //Console.WriteLine(sNewLine);
                  }
                  catch (Exception ex)
                  {
                    Console.WriteLine("Error-Comment:" + ex);
                  }
                }
              }
              #endregion

              #region code review responses
              //2. Get code review responses for this code review request            
              List<WorkItem> wicReviewResponses = GetReviewResponses(store, reviewrequest);
              if (wicReviewResponses != null && wicReviewResponses.Count > 0)
              {
                //Create new line for each code review response (CSV)
                foreach (WorkItem reviewresponse in wicReviewResponses)
                {
                  if (reviewresponse.Type.Name == "Code Review Response")
                  {
                    try
                    {
                      //"Handler", "WorkItemId", "ReviewRequestId", "ReviewResponseId", "ClosedBy", "ClosedStatus", "ClosedDate"
                      string sClosedBy = (reviewresponse.Fields["Closed By"]).Value.ToString();
                      string sClosedStatus = (reviewresponse.Fields["Closed Status"]).Value.ToString();
                      object oClosedDate = (reviewresponse.Fields["Closed Date"]).Value;
                      string sClosedDate = string.Empty;
                      if (oClosedDate != null) sClosedDate = oClosedDate.ToString();
                      string sNewLine = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}{7}", reviewrequest.CreatedBy, workitem.Id, reviewrequest.Id, reviewresponse.Id, sClosedBy, sClosedStatus, sClosedDate, Environment.NewLine);
                      swReviewResponseData.Append(sNewLine);
                      //Console.WriteLine(sNewLine);
                    }
                    catch (Exception ex)
                    {
                      Console.WriteLine("Error-CRR:" + ex);
                    }
                  }
                }
              }
              #endregion
            }
            catch (Exception ex)
            {
              Console.WriteLine("Error-WI:" + ex);
            }
          }
        }
        //});
      }

      File.WriteAllText(@"D:\Personal\Team meeting\2015\2015-Sep-All Half Yeraly Apprisal-Comments-Rev01.csv", swFileData.ToString());
      File.WriteAllText(@"D:\Personal\Team meeting\2015\2015-Sep-All Half Yeraly Apprisal-ReviewResponse-Rev01.csv", swReviewResponseData.ToString());
      File.WriteAllText(@"D:\Personal\Team meeting\2015\2015-Sep-All Half Yeraly Apprisal-ReviewRequests-Rev01.csv", swReviewRequestData.ToString());
    }

    private static string GetWorkItemLine(string sReviewRequestHeaderFormat, WorkItem workitem, List<WorkItem> wicReviewRequests)
    {
      int iReviewRequestCount = 0;
      if (wicReviewRequests != null) iReviewRequestCount = wicReviewRequests.Count;

      string sAssignedTo = workitem.Fields["Assigned To"].Value as string;
      string sTitle = workitem.Title.Replace(",", " COMMA ").Replace(Environment.NewLine, " NEWLINE ");
      string sTag = workitem.Tags.Replace(",", " COMMA ").Replace(Environment.NewLine, " NEWLINE ");

      string sEffort = string.Empty;
      if (workitem.Fields.Contains("Effort"))
      {
        object oEffort = workitem.Fields["Effort"].Value;
        if (oEffort != null) sEffort = oEffort.ToString();
      }
      object oChangedDate = workitem.Fields["Changed Date"].Value;
      string sChangedDate = string.Empty;
      if (oChangedDate != null) sChangedDate = oChangedDate.ToString();

      object oClosedDate = workitem.Fields["Closed Date"].Value;
      string sClosedDate = string.Empty;
      if (oClosedDate != null) sClosedDate = oClosedDate.ToString();

      string sRemainingWork = string.Empty;
      if (workitem.Fields.Contains("Remaining Work"))
      {
        object oRemainingWork = workitem.Fields["Remaining Work"].Value;
        if (oRemainingWork != null) sRemainingWork = oRemainingWork.ToString();
      }
      ////ID, Work Item Type, Title, Assigned To, State, Effort, Tags, Area Path, Changed Date, Closed Date, Created Date, Iteration Path, Remaining Work, ReviewRequestCount {NewLine}
      string sRequestNewLine = string.Format(sReviewRequestHeaderFormat, workitem.Id, workitem.Type.Name, sTitle, sAssignedTo, workitem.State
                                          , sEffort, sTag, workitem.AreaPath, sChangedDate, sClosedDate, workitem.CreatedDate
                                          , workitem.IterationPath, sRemainingWork, iReviewRequestCount, Environment.NewLine);
      return sRequestNewLine;
    }

    private static void AddTaskToResult(WorkItem workitem, WorkItemStore store, List<WorkItem> result)
    {
      try
      {
        string sTitle = "|" + workitem.Title.ToUpper() + "|";
        if (("|ISSUE|INTERNAL SUPPORT|INTERNALSUPPORT|").IndexOf(sTitle) >= 0)
        {
          List<WorkItem> lstTasks = GetTasks(store, workitem);
          foreach (var task in lstTasks)
            result.Add(task);
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error: Addeding Tasks" + ex);
      }
    }

    private static List<WorkItem> GetReviewResponses(WorkItemStore store, WorkItem workitem)
    {
      int id = workitem.Id;
      string queryText = "SELECT [System.Id] FROM WorkItemLinks WHERE([Target].[System.WorkItemType] = 'Code Review Response') And ([System.Links.LinkType] = 'Child') And ([Source].[System.Id] = '" + workitem.Id + "') mode(MayContain)";
      return GetChildWorkItems(store, workitem, queryText);
    }

    private static List<WorkItem> GetTasks(WorkItemStore store, WorkItem workitem)
    {      
      int id = workitem.Id;
      string queryText = "SELECT [System.Id] FROM WorkItemLinks WHERE ([Target].[System.WorkItemType] = 'Task') And ([System.Links.LinkType] = 'Child') And ([Source].[System.Id] = '" + id + "') mode(MayContain)";
      return GetChildWorkItems(store, workitem, queryText);
    }

    private static List<WorkItem> GetReviewRequests(WorkItemStore store, WorkItem workitem)
    {      
      int id = workitem.Id;
      string queryText = "SELECT [System.Id] FROM WorkItemLinks WHERE ([Target].[System.WorkItemType] = 'Code Review Request') And ([System.Links.LinkType] = 'Child') And ([Source].[System.Id] = '" + id + "') mode(MayContain)";
      return GetChildWorkItems(store, workitem, queryText);
    }

    private static List<WorkItem> GetChildWorkItems(WorkItemStore store, WorkItem workitem, string queryText)
    {
      //Ref. http://blogs.msdn.com/b/jsocha/archive/2012/02/22/retrieving-tfs-results-from-a-tree-query.aspx and
      //http://blogs.msdn.com/b/team_foundation/archive/2010/07/02/wiql-syntax-for-link-query.aspx
      //https://msdn.microsoft.com/en-us/library/bb130306%28v=vs.120%29.aspx

      List<WorkItem> details = null;
      try
      {
        int id = workitem.Id;
        var treeQuery = new Query(store, queryText);
        WorkItemLinkInfo[] links = treeQuery.RunLinkQuery();

        List<int> except = new List<int>();
        except.Add(id);

        //
        // Build the list of work items for which we want to retrieve more information
        //
        int[] ids = (from WorkItemLinkInfo info in links
                     select info.TargetId).Except(except.AsEnumerable()).ToArray();

        if (ids.Length > 0)
        {
          //
          // Next we want to create a new query that will retrieve all the column values from the original query, for
          // each of the work item IDs returned by the original query.
          //
          var detailsWiql = new StringBuilder();
          detailsWiql.AppendLine("SELECT");
          bool first = true;

          foreach (FieldDefinition field in treeQuery.DisplayFieldList)
          {
            detailsWiql.Append("    ");
            if (!first)
              detailsWiql.Append(",");
            detailsWiql.AppendLine("[" + field.ReferenceName + "]");
            first = false;
          }
          detailsWiql.AppendLine("FROM WorkItems");

          //
          // Get the work item details
          //
          var flatQuery = new Query(store, detailsWiql.ToString(), ids);
          details = flatQuery.RunQuery().OfType<WorkItem>().ToList();
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine(ex);
      }
      return details;
    }

    private static WorkItem GetWorkItemFromReviewRequest(WorkItemStore store, WorkItem workitem)
    {      
      string queryText = "SELECT [System.Id] FROM WorkItemLinks WHERE ([System.Links.LinkType] = 'Parent') And ([Source].[System.Id] = '" + workitem.Id + "') mode(MayContain)";
      return GetChildWorkItems(store, workitem, queryText)[0];
    }

    private static List<CodeReviewComment> GetCodeReviewComments(int workItemId)
    {
      List<CodeReviewComment> comments = new List<CodeReviewComment>();

      Uri uri = new Uri(URL_TO_TFS_COLLECTION);
      TeamFoundationDiscussionService service = new TeamFoundationDiscussionService();
      service.Initialize(new Microsoft.TeamFoundation.Client.TfsTeamProjectCollection(uri));
      IDiscussionManager discussionManager = service.CreateDiscussionManager();

      IAsyncResult result = discussionManager.BeginQueryByCodeReviewRequest(workItemId, QueryStoreOptions.ServerAndLocal, new AsyncCallback(CallCompletedCallback), null);
      var output = discussionManager.EndQueryByCodeReviewRequest(result);

      foreach (DiscussionThread thread in output)
      {
        if (thread.RootComment != null)
        {
          CodeReviewComment comment = new CodeReviewComment();
          comment.Author = thread.RootComment.Author.DisplayName;
          comment.Comment = thread.RootComment.Content.Replace(",", " COMMA ").Replace(System.Environment.NewLine, " NEWLINE ");
          comment.PublishDate = thread.RootComment.PublishedDate.ToShortDateString();
          comment.ItemName = thread.ItemPath;          
          comments.Add(comment);
        }
      }

      return comments;
    }

    static void CallCompletedCallback(IAsyncResult result)
    {
      // Handle error conditions here
    }

  }

  public class CodeReviewComment
  {
    public string Author { get; set; }
    public string Comment { get; set; }
    public string PublishDate { get; set; }
    public string ItemName { get; set; }
  }
}
