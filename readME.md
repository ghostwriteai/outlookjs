# Outlook JS

This library provides an unofficial interface to interact with the Outlook Web App, providing means to hook into, listen to, and interact with the app's core functionality.

## Basic Usage

First, create an instance of the class using a jQuery instance:

```javascript
const outlookJS = new outlookJS(jQuery);
```

This class can use a local jQuery instance if provided, otherwise it will try to use the global `jQuery` object. If none is found, it will attempt to load jQuery via `require('jquery')`. If all attempts fail, some functionalities may be limited.

## Key Concepts

OutlookJS exposes several ways to interact with the web app:

1. **Events:** Listen to events that occur within the web app, such as viewing an email, opening a thread, and changing recipients.

2. **DOM Methods:** Interact with the DOM directly, such as retrieving active compose windows.

3. **Fetch Observers:** Watch and respond to certain AJAX fetch requests that the web app makes, such as sending messages and fetching conversation items.

4. **Data Retrieval Methods:** Get information about the current state of the web app, like the current user's email address and current email/thread data.

5. **Cache:** The API maintains a cache of raw and processed conversation data for faster access.

## API Methods

### Event Methods

The API provides `.observe.on` and `.observe.off` methods for listening to or removing listeners for events.

It also provides `.observe.on_dom` and `.observe.off_dom` for events triggered by changes in the DOM. These methods require a specific event name and a callback function.

```javascript
outlookJs.observe.on(eventName, callback);
outlookJs.observe.off(eventName, callback);
outlookJs.observe.on_dom(eventName, callback);
outlookJs.observe.off_dom(eventName, callback);
```

### DOM Methods

The `.dom.composes()` method returns all active compose windows in the web app:

```javascript
const composeWindows = outlookJs.dom.composes();
```

### Fetch Observers

Fetch observers are used to monitor fetch requests that occur within the web app. You can use `.addFetchObserver()` to start watching for specific fetch requests:

```javascript
outlookJs.addFetchObserver(observerName, handler);
```

Currently, the API automatically observes `cacheEmailThreads` by default. Other observers can be added with the above method.

### Data Retrieval Methods

The `.get` object has several methods to retrieve the current state of the web app. These include getting the current user's sessions, email addresses, email data, thread data, and more.

```javascript
const userEmail = outlookJs.get.userEmail();
const threadData = outlookJs.get.threadData(threadId);
```

## Cache

The API maintains a cache of conversation data to improve performance and avoid unnecessary requests. This includes both raw and processed conversation data.

## Email Data Structure

An important part of interacting with the Outlook Web App is understanding the structure of the email data it provides. Here is an example of an email object that is part of a thread:

```javascript
const emailObject = {
  smtpId: email.InternetMessageId || null,
  subject: email.Subject,
  contentHtml: contentHTML,
  isDraft: email.IsDraft || false,
  from: from,
  to: to,
  cc: cc,
  bcc: bcc,
  createdAt: email.DateTimeReceived || 0,
  sentAt: email.DateTimeSent || 0,
  threadId: email.ConversationId?.Id || this.getSelectedConversationId(),
  operation: email.operation || null,
  timestamp: timestamp,
};
```

**Important Note:** Outlook does not associate a unique, static ID with each email message. It's up you how to implement IDs as you see appropriate.

## Message Parser from DOM

The `parseThreadFromDom` method is a utility that parses information from an opened email thread in the DOM. This method returns an object that includes a list of email objects. Each email object contains a randomly generated ID and data about the email, including sender, recipients, and content.

```javascript
const threadData = outlookJs.parseThreadFromDom();
```

**Important Note:** The `parseThreadFromDom` method can only retrieve information from opened emails and the preview text from collapsed ones.

## Notes

This API is still under development and may change in the future. Always ensure to update your codebase according to the latest changes in the API.

## Future Development

The plan is to extend the API with more convenience methods and data processing features. Contributions are welcome, please follow the contribution guidelines.

## Fetch Observer Event Names

Currently, the following fetch request event names are available for observation:

1. **sendMessage**: This event fires when a new email message is being sent. The associated requests are `createitem` and `updateitem`.
2. **cacheEmailThreads**: This event fires when email thread data is being fetched. The associated request is `getconversationitems`.

To observe these events, use the `.addFetchObserver()` method, providing the event name and a callback function:

```javascript
outlookJs.addFetchObserver(eventName, callback);
```

## DOM Observer Event Names

The current DOM event names are as follows:

1. **viewEmail**: Triggered when an email is viewed. The associated class name for this event is `.SlLx9.byzS1`.
2. **viewThread**: Triggered when a thread is opened. The associated class name is also `.SlLx9.byzS1`.
3. **compose**: Triggered when a new email compose window is opened. The associated class names are `.dMm6A` and `.gXGox`.
4. **recipientChange**: Triggered when the recipient of an email is changed. The associated class name is `.Lbs4W`.

To listen to these events, use the `.observe.on_dom()` method, providing the event name and a callback function:

```javascript
outlookJs.observe.on_dom(eventName, callback);
```

To stop listening to these events, use the `.observe.off_dom()` method, again providing the event name and the callback function:

```javascript
outlookJs.observe.off_dom(eventName, callback);
```

## Important Note

Please be aware that these event names are tied to class names that Outlook Web App uses internally. If Outlook changes these class names in the future, these events may no longer work as expected. Keep your application updated with any changes to this API to maintain compatibility.
