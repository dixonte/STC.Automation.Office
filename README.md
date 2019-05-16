# STC.Automation.Office
A small, incomplete C# library that provides late-binding wrappers for dealing with Office products in a version-agnostic fashion. Implements IDisposable to release COM objects, and provides some other helpful features.

## Contributing

There are two primary resources to assist development:

1. The online documentation of the Office application object models, found here: https://docs.microsoft.com/en-us/office/vba/api/overview/
2. Oleview.exe to inspect the COM interfaces of the Office libraries to obtain the interface ID's. You can get a copy of oleview.exe in the Windows SDK's. For Windows 10, you can download the SDK here: https://developer.microsoft.com/en-US/windows/downloads/windows-10-sdk

### Example

```csharp
/// <summary>
/// Represents a document or link to a document contained in an Outlook item.
/// </summary>
[WrapsCOM("Outlook.Attachment", "00063007-0000-0000-C000-000000000046")]
public class Attachment : ComWrapper
{
    internal Attachment(object attachmentObj)
        : base(attachmentObj)
    {
    }
}
```

Where ```attachmentObj``` is the COM object for an attachment in Outlook, which has interface ID ```00063007-0000-0000-C000-000000000046``` as exposed in ```oleview.exe``` for type library ```"C:\Program Files\Microsoft Office\root\Office16\MSOUTL.OLB"```

An ```Attachment``` object can be created by indexing an ```Attachments``` object as follows (error handling omitted for clarity):

```csharp
/// <summary>
/// Returns an Attachment object from the collection.
/// </summary>
public Attachment this[int key]
    => new Attachment(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
```

Where ```InternalObject``` is the COM object for the attachments collection in Outlook. ```InvokeMember``` calls a property or method on the COM object in a late-bound fashion, taking an untyped array of parameters matching the [method signature found in the documentation](https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments.item).

*Note that the Attachment class inherits from ComWrapper which is an IDisposable object so consumers must dispose of the attachment object when they're finished.*

For convenience, COM objects are sometimes cached in their consuming class such that they're disposed within the same lifetime. In the following example, the reference to the Attachments collection on an Outlook MailItem is internally cached and disposed when the MailItem is disposed.

```csharp
/// <summary>
/// Wraps an Outlook.MailItem object
/// </summary>
[WrapsCOM("Outlook.MailItem", "00063034-0000-0000-C000-000000000046")]
public class MailItem : OutlookItem
{
    private Attachments _attachments;

    /// <summary>
    /// Returns an Attachments object that represents all the attachments for the specified item. This object is internally cached and does not require manual disposal.
    /// </summary>
    public Attachments Attachments
    {
        get
        {
            if (_attachments == null)
                _attachments = new Attachments(InternalObject.GetType().InvokeMember("Attachments", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

            return _attachments;
        }
    }

    
    internal override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // Free managed
            if (_attachments != null)
            {
                _attachments.Dispose();
                _attachments = null;
            }
        }

        base.Dispose(true);
    }
}
```