using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook.Enums
{
    /// <summary>
    /// Specifies the folder type for a specified folder.
    /// </summary>
    public enum DefaultFolders
    {
        /// <summary>
        /// The Calendar folder.
        /// </summary>
        Calendar = 9,
        /// <summary>
        /// The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
        /// </summary>
        Conflicts = 19,
        /// <summary>
        /// The Contacts folder.
        /// </summary>
        Contacts = 10,
        /// <summary>
        /// The Deleted Items folder.
        /// </summary>
        DeletedItems = 3,
        /// <summary>
        /// The Drafts folder.
        /// </summary>
        Drafts = 16,
        /// <summary>
        /// The Inbox folder.
        /// </summary>
        Inbox = 6,
        /// <summary>
        /// The Journal folder.
        /// </summary>
        Journal = 11,
        /// <summary>
        /// The Junk E-Mail folder.
        /// </summary>
        Junk = 23,
        /// <summary>
        /// The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
        /// </summary>
        LocalFailures = 21,
        /// <summary>
        /// The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
        /// </summary>
        ManagedEmail = 29,
        /// <summary>
        /// The Notes folder.
        /// </summary>
        Notes = 12,
        /// <summary>
        /// The Outbox folder.
        /// </summary>
        Outbox = 4,
        /// <summary>
        /// The Sent Mail folder.
        /// </summary>
        SentMail = 5,
        /// <summary>
        /// The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
        /// </summary>
        ServerFailures = 22,
        /// <summary>
        /// The Suggested Contacts folder.
        /// </summary>
        SuggestedContacts = 30,
        /// <summary>
        /// The Sync Issues folder. Only available for an Exchange account.
        /// </summary>
        SyncIssues = 20,
        /// <summary>
        /// The Tasks folder.
        /// </summary>
        Tasks = 13,
        /// <summary>
        /// The To Do folder.
        /// </summary>
        ToDo = 28,
        /// <summary>
        /// The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
        /// </summary>
        PublicFoldersAllPublicFolders = 18,
        /// <summary>
        /// The RSS Feeds folder.
        /// </summary>
        RssFeeds = 25
    }
}
