#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using microsoft.graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.OData.Client;
using Microsoft.OData.ProxyExtensions;
using System.Net.Http;

#endregion

namespace MicrosoftGraphSampleConsole
{
    internal class Requests
    {
        public static microsoft.graph.GraphService client;
        public static void UserMode()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            #region Setup Microsoft Graph Client for user

            //*********************************************************************
            // setup Microsoft Graph Client for user...
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsUser();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }

            #endregion

            Console.WriteLine("\nStarting user-mode requests...");
            Console.WriteLine("\n=============================\n\n");

            #region Get the signed in user's details, manager, reports and group memberships
            //// GET /me

            user user = new user();
            try
            {

                user = (user)client.me.ExecuteAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me");
                Console.WriteLine();
                Console.WriteLine("    Id: {0}  UPN: {1}", user.id, user.userPrincipalName);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting /me user {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            try
            {
                // get signed in user's picture.  
                // Drop to REST for this - Client Library doesn't support this yet :(
                string token = AuthenticationHelper.TokenForUser;
                string request = "me/Photo/$value";
                Stream photoStream = Helper.GetRestRequestStream(request, token).Result;
                Console.WriteLine("Got stream photo");
            }
            catch (Exception)
            {
                Console.WriteLine("Failed to get stream");
            }
             
            try
            {
                // GET /me/directReports
                List<IdirectoryObject> directs = client.me.directReports.ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/directReports");
                Console.WriteLine();
                if (directs.Count == 0)
                {
                    Console.WriteLine("      no reports");
                }
                else
                {
                    foreach (IdirectoryObject _user in directs)
                    {
                        if (_user is Iuser)
                        {
                            Iuser __user = (Iuser)_user;
                            Console.WriteLine("      Id: {0}  UPN: {1}", __user.id, __user.userPrincipalName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting directReports {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            try
            {
                // GET /me/manager
                IdirectoryObject manager = client.me.manager.ExecuteAsync().Result;
                Console.WriteLine();
                Console.WriteLine("GET /me/manager");
                Console.WriteLine();
                if (manager == null)
                {
                    Console.WriteLine("      no manager");
                }
                else
                {
                    Iuser _user = client.users.GetById(manager.id).ExecuteAsync().Result;
                    Iuser __user = (Iuser)_user;
                    Console.WriteLine("\nManager      Id: {0}  UPN: {1}", __user.id, __user.userPrincipalName);
                    //    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting directReports {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            // GET /me/memberOf
            try
            {
                IuserFetcher uFetcher = user;
                List<IdirectoryObject> _groups = uFetcher.memberOf.ExecuteAsync().Result.CurrentPage.ToList();
                // List<IdirectoryObject> _groups = client.me.memberOf.ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/memberOf");
                Console.WriteLine();
                if (_groups.Count == 0)
                {
                    Console.WriteLine("    user is not a member of any groups");
                }
                foreach (IdirectoryObject _group in _groups)
                {
                    if (_group is group)
                    {
                        group __group = _group as group;
                        Console.WriteLine("    Id: {0}  UPN: {1}", __group.id, __group.displayName);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting group memberships {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            #endregion

            #region Get the signed in user's files, who last modified them, messages and events, and personal contacts
            try
            {
                IList<IdriveItem> _items = client.me.drive.root.children.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/files?$top=5");
                Console.WriteLine();
                foreach (IdriveItem _item in _items)
                {

                    if (_item.file != null)
                    {
                        Console.WriteLine("    This is a folder: File Id: {0}  WebUrl: {1}", _item.id, _item.webUrl);
                    }
                    else
                    {

                        Console.WriteLine("    File Id: {0}  WebUrl: {1}", _item.id, _item.webUrl);
                    }
                }
            }

            catch (Exception e)
            {
                Console.WriteLine("\nError getting files or content {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }

            try
            {

                // GET /me/Messages?$top=5
                List<Imessage> messages = client.me.messages.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/messages?$top=5");
                Console.WriteLine();
                if (messages.Count == 0)
                {
                    Console.WriteLine("    no messages in mailbox");
                }
                foreach (Imessage message in messages)
                {
                    Console.WriteLine("    Message: {0} received {1} ", message.subject, message.receivedDateTime);
                }

                // GET /me/Events?$top=5
                List<IEvent> events = client.me.events.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/events?$top=5");
                Console.WriteLine();
                if (events.Count == 0)
                {
                    Console.WriteLine("    no events scheduled");
                }
                foreach (IEvent _event in events)
                {
                    Console.WriteLine("    Event: {0} starts {1} ", _event.subject, _event.start);
                }

                // GET /me/contacts?$top=5
                List<Icontact> myContacts = client.me.contacts.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /me/myContacts?$top=5");
                Console.WriteLine();
                if (myContacts.Count == 0)
                {
                    Console.WriteLine("    You don't have any contacts");
                }
                foreach (Icontact _contact in myContacts)
                {
                    Console.WriteLine("    Contact: {0} ", _contact.displayName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting messages, events and contacts {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            #endregion

            #region People picker example
            //*********************************************************************
            // People picker
            // Search for a user using text string "Sample" match against userPrincipalName, displayName, giveName, surname
            // Change searchString to suite your needs
            //*********************************************************************
            Console.WriteLine("\nSearch for user (enter search string):");
            String searchString = Console.ReadLine();

            List<Iuser> usersList = null;
            IPagedCollection<Iuser> searchResults = null;
            try
            {
                IuserCollection userCollection = client.users;
                searchResults = userCollection.Where(u =>
                    u.userPrincipalName.StartsWith(searchString) ||
                    u.displayName.StartsWith(searchString) ||
                    u.givenName.StartsWith(searchString) ||
                    u.surname.StartsWith(searchString)).Take(5).ExecuteAsync().Result;
                usersList = searchResults.CurrentPage.ToList();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting User {0} {1}", e.Message,
                    e.InnerException != null ? e.InnerException.Message : "");
            }

            if (usersList != null && usersList.Count > 0)
            {
                do
                {
                    int j = 1;
                    usersList = searchResults.CurrentPage.ToList();
                    foreach (Iuser u in usersList)
                    {
                        Console.WriteLine("User {2} DisplayName: {0}  UPN: {1}",
                            u.displayName, u.userPrincipalName, j);
                        j++;
                    }
                    searchResults = searchResults.GetNextPageAsync().Result;
                } while (searchResults != null);
            }
            else
            {
                Console.WriteLine("User not found");
            }

            #endregion

            #region Create a unified group
            // POST /groups - create unified group 
            Console.WriteLine("\nDo you want to create a new unified group? Click y/n\n");
            ConsoleKeyInfo key = Console.ReadKey();
            group uGroup = null;
            if (key.KeyChar == 'y')
            {
                string suffix = Helper.GetRandomString(5);
                uGroup = new group
                {
                    groupTypes = new List<string> { "Unified" },
                    displayName = "Unified group " + suffix,
                    description = "Group " + suffix + " is the best ever",
                    mailNickname = "Group" + suffix,
                    mailEnabled = true,
                    securityEnabled = false
                };
                try
                {
                    client.groups.AddgroupAsync(uGroup).Wait();
                    Console.WriteLine("\nCreated unified group {0}", uGroup.displayName);
                }
                catch (Exception)
                {
                    Console.WriteLine("\nIssue creating the group {0}", uGroup.displayName);
                    uGroup = null;
                }
            }

            #endregion

            #region Add/remove group members
            // Currently busted through the client library.  Investigating...  Will re-add commented section once fixed.
            client.Context.SaveChanges();
            group retreivedGroup = new group();
            // get a set of users to add
            List<Iuser> members = client.users.Take(3).ExecuteAsync().Result.CurrentPage.ToList();

            // Either add to newly created group, OR add to an existing group
            if (uGroup != null)
            {
                retreivedGroup = (group)client.groups.GetById(uGroup.id).ExecuteAsync().Result;
            }
            else
            {
                List<Igroup> foundGroups = client.groups.Where(ug => ug.groupTypes.Any(gt => gt == "Unified")).Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                if (foundGroups != null && foundGroups.Count > 0)
                {
                    retreivedGroup = foundGroups.First() as group;
                }
            }

            retreivedGroup.SaveChangesAsync().Wait();

            //if (retreivedGroup != null)
            //{
            //    // Add users
            //    foreach (Iuser _user in members)
            //    {
            //        try
            //        {
            //            directoryObject user1 = new user();
            //            user1.id = _user.id;
            //            retreivedGroup.members.Add(user1);
            //            retreivedGroup.UpdateAsync().Wait();
            //            Console.WriteLine("\nAdding {0} to group {1}", _user.userPrincipalName, uGroup.displayName);
            //        }
            //        catch (Exception e)
            //        {
            //            Console.WriteLine("\nError assigning member to group. {0} {1}",
            //                 e.Message, e.InnerException != null ? e.InnerException.Message : "");
            //        }
            //    }

            //    // now remove the added users
            //    foreach (user _user in members)
            //    {
            //        try
            //        {
            //            retreivedGroup.members.Remove(_user as directoryObject);
            //            retreivedGroup.UpdateAsync().Wait();
            //            Console.WriteLine("\nRemoved {0} from group {1}", _user.userPrincipalName, retreivedGroup.displayName);
            //        }
            //        catch (Exception e)
            //        {
            //            Console.WriteLine("\nError removing member from group. {0} {1}",
            //                 e.Message, e.InnerException != null ? e.InnerException.Message : "");
            //        }
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("\nCan't find any unified groups to add members to.\n");
            //}

            #endregion

            #region Get groups
            // GET /groups?$top=5
            List<Igroup> groups = client.groups.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
            Console.WriteLine();
            Console.WriteLine("GET /groups?$top=5");
            Console.WriteLine();
            foreach (Igroup _group in groups)
            {
                Console.WriteLine("    Group Id: {0}  upn: {1}", _group.id, _group.displayName);
                foreach (string _type in _group.groupTypes)
                {
                    if (_type == "Unified")
                    {
                        Console.WriteLine(": This is a Unifed Group");
                    }
                }
            }
            #endregion

            #region Get the first 3 UNIFIED groups and view their associated content
            // GET /groups?$top=5&$filter=groupType eq 'Unified'
            // groups = client.me.memberOf.OfType<Igroup>().Where(ug => ug.groupTypes.Any(gt => gt == "Unified")).Take(3).ExecuteAsync().Result.CurrentPage.ToList();
            groups = client.groups.Where(ug => ug.groupTypes.Any(gt => gt == "Unified")).Take(3).ExecuteAsync().Result.CurrentPage.ToList();
            Console.WriteLine();
            Console.WriteLine("GET /groups?$top=5&$filter=groupType eq 'Unified'");
            Console.WriteLine();
            foreach (Igroup _group in groups)
            {
                Console.WriteLine("    Unified Group: {0}", _group.displayName);

                try
                {
                    // get group members
                    List<IdirectoryObject> unifiedGroupMembers = client.groups.GetById(_group.id).members.ExecuteAsync().Result.CurrentPage.ToList();
                    if (unifiedGroupMembers.Count == 0)
                    {
                        Console.WriteLine("      no members for group");
                    }
                    foreach (IdirectoryObject _member in unifiedGroupMembers)
                    {
                        if (_member is Iuser)
                        {
                            Iuser memberUser = (Iuser)_member;
                            Console.WriteLine("        User: {0} ", memberUser.displayName);
                        }
                    }

                    //get group files
                    try
                    {
                        IList<IdriveItem> unifiedGroupFiles = client.groups.GetById(_group.id).drive.root.children.Take(5).ExecuteAsync().Result.CurrentPage.ToList();
                        if (unifiedGroupFiles.Count == 0)
                        {
                            Console.WriteLine("      no files for group");
                        }
                        foreach (IdriveItem _file in unifiedGroupFiles)
                        {
                            Console.WriteLine("        file: {0} url: {1}", _file.name, _file.webUrl);
                        }
                    }
                    catch (Exception)
                    {
                        Console.Write("Unexpected exception when enumerating group files");
                    }

                    //get group conversations
                    try
                    {
                        List<Iconversation> unifiedGroupConversations = client.groups.GetById(_group.id).conversations.ExecuteAsync().Result.CurrentPage.ToList();
                        if (unifiedGroupConversations.Count == 0)
                        {
                            Console.WriteLine("      no conversations for group");
                        }
                        foreach (Iconversation _conversation in unifiedGroupConversations)
                        {
                            Console.WriteLine("        conversation topic: {0} ", _conversation.topic);
                        }
                    }
                    catch (Exception)
                    {
                        Console.Write("Unexpected exception when enumerating group conversations");
                    }

                    //get group events
                    try
                    {
                        List<IEvent> unifiedGroupEvents = client.groups.GetById(_group.id).events.ExecuteAsync().Result.CurrentPage.ToList();
                        if (unifiedGroupEvents.Count == 0)
                        {
                            Console.WriteLine("      no meeting events for group");
                        }
                        foreach (IEvent _event in unifiedGroupEvents)
                        {
                            Console.WriteLine("        meeting event subject: {0} ", _event.subject);
                        }
                    }
                    catch (Exception)
                    {
                        Console.Write("Unexpected exception when enumerating group events");
                    }
                }
                catch (Exception)
                {
                    Console.Write("Unexpected exception when enumerating group members and events");
                }
            }
            #endregion

            #region Get the top 10 users and create a recipient list (to be used later)
            IList<recipient> messageToList = new List<recipient>();

            // GET /users?$top=5
            List<Iuser> users = client.users.Take(10).ExecuteAsync().Result.CurrentPage.ToList();
            foreach (Iuser _user in users)
            {
                if (_user.assignedPlans.Count != 0)
                {
                    recipient messageTo = new recipient();
                    emailAddress emailAdress = new emailAddress();
                    emailAdress.address = _user.userPrincipalName;
                    emailAdress.name = _user.displayName;
                    messageTo.emailAddress = emailAdress;
                    messageToList.Add(messageTo);
                }
            }

            // also add current signed in user to the recipient list IF they have a license
            if (user.assignedPlans.Count != 0)
            {
                recipient messageTo = new recipient();
                emailAddress emailAdress = new emailAddress();
                emailAdress.address = user.userPrincipalName;
                emailAdress.name = user.displayName;
                messageTo.emailAddress = emailAdress;
                messageToList.Add(messageTo);
            }
            #endregion

            #region Send mail to signed in user and the recipient list
            // POST /me/SendMail

            Console.WriteLine();
            Console.WriteLine("POST /me/sendmail");
            Console.WriteLine();

            try
            {
                itemBody messageBody = new itemBody();
                messageBody.content = "<report pending>";
                messageBody.contentType = bodyType.text;

                message newMessage = new message();
                newMessage.subject = string.Format("\nCompleted test run from console app at {0}.", currentDateTime);
                newMessage.toRecipients = (IList<recipient>)messageToList;
                newMessage.body = (itemBody)messageBody;

                client.me.sendMailAsync(newMessage, true);

                Console.WriteLine("\nMail sent to {0}", user.displayName);
            }
            catch (Exception)
            {
                Console.WriteLine("\nUnexpected Error attempting to send an email");
                throw;
            }
            #endregion

            #region clean up (delete any created items)
            if (uGroup != null)
            {
                try
                {
                    uGroup.DeleteAsync().Wait();
                    Console.WriteLine("\nDeleted group {0}", uGroup.displayName);
                }
                catch (Exception e)
                {
                    Console.Write("Couldn't delete group.  Error detail: {0}", e.InnerException.Message);
                }
            }
            #endregion

        }
        public static void AppMode()
        {
            // record start DateTime of execution
            string currentDateTime = DateTime.Now.ToUniversalTime().ToString();
            #region Setup Microsoft Graph Client for app
            //*********************************************************************
            // setup Microsoft Graph Client for app
            //*********************************************************************
            try
            {
                client = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Acquiring a token failed with the following error: {0}", ex.Message);
                if (ex.InnerException != null)
                {
                    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                    //InnerException Message will contain the HTTP error status codes mentioned in the link above
                    Console.WriteLine("Error detail: {0}", ex.InnerException.Message);
                }
                Console.ResetColor();
                Console.ReadKey();
                return;
            }
            #endregion

            Console.WriteLine("\nStarting app-mode requests...");
            Console.WriteLine("\n=============================\n\n");

            #region get specific user's email messages
            // 
            // get a specific user's email - demonstrates $select in Linq query
            try
            {
                Console.WriteLine("\nEnter the email address of the user mailbox you want to retrieve:");
                String email = Console.ReadLine();
                List<Imessage> messages = client.users.GetById(email).messages.Select<Imessage>(y => new message(){subject = y.subject, receivedDateTime = y.receivedDateTime}).Take(3).
                    ExecuteAsync().Result.CurrentPage.ToList();
                Console.WriteLine();
                Console.WriteLine("GET /user/{0}/messages?$top=3&$select=subject,receivedDateTime", email);
                Console.WriteLine();
                if (messages.Count == 0)
                {
                    Console.WriteLine("    no messages in mailbox");
                }
                foreach (Imessage message in messages)
                {
                    Console.WriteLine("    Message: {0} received {1} ", message.subject, message.receivedDateTime);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nError getting files or content {0} {1}",
                     e.Message, e.InnerException != null ? e.InnerException.Message : "");
            }
            #endregion
        }
    }
}