export class outlookJS {
  constructor(localJQuery) {
    this.callbacks = new Map();
    this.cache = {};
    this.cache.rawConversationMap = new Map();
    this.cache.processedConversationMap = new Map();
    this.observer = this.createObserver();
    this.fetchObservers = {
      requestObservers: [],
      responseObservers: [],
    };
    this.addFetchObserver("cacheEmailThreads", this.cacheEmailThreads); // Add observer by default to start caching threads
    this.initFetchWatcher();

    var $;
    if (typeof localJQuery !== "undefined") {
      $ = localJQuery;
    } else if (typeof jQuery !== "undefined") {
      $ = jQuery;
    } else {
      // try load jQuery through node.
      try {
        $ = require("jquery");
      } catch (err) {
        // else leave $ undefined, which may be fine for some purposes.
      }
    }

    this.observe = {
      on: (eventName, callback) => {
        this.on(eventName, callback);
      },
      off: (eventName, callback) => {
        this.off(eventName, callback);
      },
      on_dom: (eventName, callback) => {
        const className = this.getClassNameForEvent(eventName);
        if (className) {
          this.on_dom(className, callback);
        }
      },
      off_dom: (eventName, callback) => {
        const className = this.getClassNameForEvent(eventName);
        if (className) {
          this.off_dom(className, callback);
        }
      },
    };

    this.get = {
      userSessions: () => this.getUserSessions(),
      currentEmailAddress: () => this.getCurrentEmailAddress(),
      userEmail: () => this.getCurrentEmailAddress(),
      allEmailAddresses: () => this.getAllEmailAddresses(),
      emailData: () => this.getEmailData(),
      emailId: () => this.getEmailId(),
      threadData: (threadId) => this.getThread(threadId),
      threadId: () => this.getSelectedConversationId(),
      currentEmail: (thread, emailId) => {
        // again both params to standardize AutoReply

        return thread?.emails && thread.emails.length > 0
          ? thread.emails[thread.emails.length - 1]
          : null;
      },
      emailsWithoutDrafts: (thread) =>
        thread["emails"].filter((email) => !email.is_draft),
      emailReplyButton: () => {
        return $(this.selectors.responseButton).first();
      },
      backupEmailReplyButton: () => this.get.emailReplyButton(),
      threadIdFromEmail: (emailId) => this.getSelectedConversationId(),
    };

    this.check = {
      isInsideEmail: () => this.isInsideEmail(),
    };

    this.dom = {
      composes: () => this.getComposeWindows(),
    };
  }

  getComposeWindows() {
    let composeElements = [];

    // hardcoded need to change this is real compose window
    $(".dMm6A").each((i, el) => {
      composeElements.push(new Compose(el));
    });

    return composeElements;
  }

  getClassNameForEvent(eventName) {
    const eventClassMap = {
      viewEmail: ".SlLx9.byzS1",
      viewThread: ".SlLx9.byzS1",
      compose: ".dMm6A, .gXGox",
      recipientChange: ".Lbs4W",
    };

    return eventClassMap[eventName];
  }

  // object that stores ConversationNodes which form an email thread
  // still used if only singular email
  getConversation(response) {
    return response.Body.ResponseMessages.Items[0].Conversation;
  }

  getThreadId(conversation) {
    return conversation.ConversationId.Id;
  }

  convertDateToUnix(date) {
    const jsDate = new Date(date);
    const unixTime = Math.floor(jsDate.getTime() / 1000);
    return unixTime;
  }

  getEmailData(conversationNode) {
    const email = conversationNode.Items[0];
    const emptyContent = "<html><body></body></html>";

    if (!email.UniqueBody && !email.NewBodyContent) return; // Ignore emails with no body
    if (email.UniqueBody && email.UniqueBody.Value === emptyContent) return; // Ignore emails with empty body

    const recipientMapper = (recipient) => ({
      address: recipient.EmailAddress,
      name: recipient.Name,
    });

    try {
      const contentHTML = email.UniqueBody
        ? email.UniqueBody.Value
        : email.NewBodyContent.Value;
      const from = email.From.Mailbox
        ? {
            address: email.From.Mailbox.EmailAddress,
            name: email.From.Mailbox.Name,
          }
        : { address: this.getCurrentEmailAddress(), name: "" };
      const to = email.ToRecipients?.map(recipientMapper) || [];
      const cc = email.CcRecipients?.map(recipientMapper) || [];
      const bcc = email.BccRecipients?.map(recipientMapper) || [];
      const timestamp = email.DateTimeSent
        ? this.convertDateToUnix(email.DateTimeSent)
        : 0;

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

      return emailObject;
    } catch (e) {
      return;
    }
  }

  getThreadDataWithFallback(existingThreadId) {
    if (!existingThreadId) existingThreadId = "";
    let threadData = this.get.threadData(existingThreadId);

    // this can be null, which means we want to trigger a parse within the construction to make sure the DOM is fully loaded
    return threadData;
  }

  addFetchObserver(fetchObserver, handler) {
    const requestObservers = {
      sendMessage: ["createitem", "updateitem"],
    };

    const responseObservers = {
      cacheEmailThreads: ["getconversationitems"],
    };

    const observer = {
      urlPatterns:
        requestObservers[fetchObserver] || responseObservers[fetchObserver],
      handler,
    };

    if (requestObservers[fetchObserver]) {
      this.fetchObservers.requestObservers.push(observer);
    } else if (responseObservers[fetchObserver]) {
      this.fetchObservers.responseObservers.push(observer);
    } else {
      console.error(`Invalid fetchObserver: ${fetchObserver}`);
    }
  }

  handleObserverResponse(url, clonedResponse) {
    clonedResponse
      .json()
      .then((json) => {
        this.fetchObservers.responseObservers.forEach((observer) => {
          observer.urlPatterns.forEach((urlPattern) => {
            if (url.includes(urlPattern)) {
              observer.handler(json);
            }
          });
        });
      })
      .catch((e) => {});
  }

  handleObserverRequest(url, args) {
    this.fetchObservers.requestObservers.forEach((observer) => {
      observer.urlPatterns.forEach((urlPattern) => {
        if (url.includes(urlPattern)) {
          observer.handler(args);
        }
      });
    });
  }

  initFetchWatcher() {
    if (this.fetchInitialized) {
      return;
    }
    this.fetchInitialized = true;
    const win = top;

    const originalFetch = win.fetch;

    win.fetch = new Proxy(originalFetch, {
      apply: async (target, thisArg, args) => {
        const url = args[0].toLowerCase();

        try {
          this.handleObserverRequest(url, args);
        } catch (error) {
          // console.error(error);
        }

        if (url.includes("getconversationitems")) {
          if (!this.conversationRequestOptions) {
            this.conversationRequestOptions = args[1];
            let newBody = JSON.parse(this.conversationRequestOptions.body);
            newBody.Body.SyncState = ""; // clear sync state to get all emails
            this.conversationRequestOptions.body = JSON.stringify(newBody);
            args[1] = this.conversationRequestOptions;
          }
        }

        const response = await Reflect.apply(target, thisArg, args);

        try {
          this.handleObserverResponse(url, response.clone());
        } catch (error) {
          // console.error(error)
        }

        return response;
      },
    });
  }

  async sendConversationRequest(conversationId) {
    let newBody = JSON.parse(this.conversationRequestOptions.body);
    newBody.Body.Conversations[0].ConversationId.Id = conversationId;
    let newOptions = this.conversationRequestOptions;
    newOptions.body = JSON.stringify(newBody);

    let res = await window.fetch(
      "/owa/0/service.svc?action=GetConversationItems&app=Mail&n=1",
      newOptions
    );
    let json = await res.json();
    return json;
  }

  cacheEmailThreads = (response) => {
    const conversation = this.getConversation(response);
    const conversationId = conversation.ConversationId.Id;
    const rawMap = this.cache.rawConversationMap;
    const processedMap = this.cache.processedConversationMap;

    if (conversation.ConversationNodes && !rawMap.has(conversationId)) {
      rawMap.set(conversationId, conversation);

      const emailDataList = conversation.ConversationNodes.map((node) => {
        return this.getEmailData(node);
      }).filter((emailData) => emailData !== undefined);

      if (emailDataList.length > 0) {
        if (processedMap.has(conversationId)) {
          const existingList = processedMap.get(conversationId);
          if (emailDataList.length > existingList.length) {
            processedMap.set(conversationId, {
              emails: emailDataList,
              thread_id: conversationId,
            }); // naming bc of gmailjs
          }
        } else {
          processedMap.set(conversationId, {
            emails: emailDataList,
            thread_id: conversationId,
          }); // naming bc of gmailjs
        }
      }
    }
  };

  getAllEmails = () => {
    return this.cache.processedConversationMap;
  };

  getThread(threadId) {
    if (!threadId) threadId = this.get.threadId();
    const conversation = this.cache.processedConversationMap.get(threadId);
    if (conversation) {
      return conversation;
    }
    return null;
  }

  createObserver() {
    return new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        if (mutation.type === "childList") {
          const addedNodes = Array.from(mutation.addedNodes);

          for (const node of addedNodes) {
            if (node.nodeType === Node.ELEMENT_NODE) {
              this.callbacks.forEach((callbacks, className) => {
                const foundNode = $(node)
                  .find(className)
                  .add($(node).filter(className));
                if (foundNode.length > 0) {
                  // only thing currently tracked by library is compose windows
                  callbacks.forEach((callback) =>
                    callback(new Compose(foundNode))
                  );
                }
              });
            }
          }
        }
      }
    });
  }

  on(requestParamList, callback) {
    this.addFetchObserver(requestParamList, callback);
  }

  off(requestParamList, callback) {
    this.fetchObservers = this.fetchObservers.filter((observer) => {
      return observer.urlPatterns !== requestParamList;
    });
  }

  on_dom(className, callback) {
    // Set up the observer to observe the entire document
    if (this.callbacks.size === 0) {
      this.observer.observe(document.body, {
        childList: true,
        subtree: true,
      });
    }

    // Register the callback for the specified className
    if (!this.callbacks.has(className)) {
      this.callbacks.set(className, []);
    }
    this.callbacks.get(className).push(callback);

    // this.checkExistingNodes(className)
  }

  checkExistingNodes(className) {
    const existingNodes = document.querySelectorAll(className);
    if (existingNodes.length > 0) {
      const jQueryNodes = $(existingNodes);
      this.callbacks
        .get(className)
        .forEach((callback) => callback(jQueryNodes));
    }
  }

  off_dom(className, callback) {
    const registeredCallbacks = this.callbacks.get(className);
    if (registeredCallbacks) {
      const index = registeredCallbacks.indexOf(callback);
      if (index !== -1) {
        registeredCallbacks.splice(index, 1);
      }
      // If there are no callbacks left for the specified className, disconnect the observer
      if (registeredCallbacks.length === 0) {
        this.callbacks.delete(className);
        if (this.callbacks.size === 0) {
          this.observer.disconnect();
        }
      }
    }
  }

  // retrieves current session ids and times
  getUserSessions() {
    const allSessionsObject = JSON.parse(
      window.localStorage.sessionTracking_SignedInAccountList
    );
    return allSessionsObject;
  }

  getCurrentEmailAddress() {
    const sessionIds = this.getUserSessions(window);

    // find the largest lastActiveTime value in a list of objects and return the object with the largest value
    const mostRecentSession = sessionIds.reduce((a, b) =>
      a.lastActiveTime > b.lastActiveTime ? a : b
    );

    const currentSession = JSON.parse(
      window.localStorage[mostRecentSession.sessionTrackingKey]
    );
    const userEmail = currentSession.upn;

    return userEmail;
  }

  getAllEmailAddresses() {
    const sessionIds = this.getUserSessions(window);

    const emails = sessionIds.map((session) => {
      const currentSession = JSON.parse(
        window.localStorage[session.sessionTrackingKey]
      );
      const userEmail = currentSession.upn;
      return userEmail;
    });

    return emails;
  }

  // checks if the current page is an email
  isInsideEmail() {
    // check if the current window contains /id/ in the url
    return window.location.href.includes("/id/"); // hardcoded need to change
  }

  // retrieves the conversation id from the current page
  getSelectedConversationId() {
    const url = window.location.href;
    try {
      let conversationId = decodeURIComponent(url.split("/id/")[1]);
      return conversationId;
    } catch (e) {
      // no id return empty string
      return "";
    }
  }

  selectors = {
    bodyDiv: ".dFCbN.dPKNh.DziEn", // first child of this class
    recipientInput: ".vBoqL",
    lastGu: ".f1MMn",
    leftBar: ".___1cp4gt1",
    emailBody: ".aVla3",
    emailThread: ".Q8TCC.yyYQP.customScrollBar",
    autoReply: ".cIAN3.DNevi",
    autoReplyContainer: ".LIQKd.G_Fzw",
    responseButtonContainer: ".th6py",
    subject: ".f1MMn",
    lowestComposeInput: ".yz4r1", // lowestComposeInput differs based on whether it's a compose (subject) or reply (recipientInput)
    replyParent: "#editorParent_1",
    sendButton: ".b56tW",
    responseButton: ".xgQkx.FGlQz",
    composeButton: "[aria-label='New mail']",
    composeWindow:
      ".soZTT.UoDR_.customScrollBar.Qmg2Q, .Q8TCC.yyYQP.customScrollBar", // first is for compose, second is for reply .Q8TCC.yyYQP.customScrollBar
    replyThread: ".aVla3.gXGox",
    messageBody: ".dFCbN.k1Ttj.dPKNh.DziEn, [id^=editorParent_]",
    subject: `input[aria-label="Add a subject"]`,
    threadSubject: ".full.UAxMv",
    threadMessage: ".aVla3",
    threadSender: ".OZZZK",
    recipientContainer: ".ow896",
    recipient: "div.IIvzX",
    threadBody: ".XbIp4.jmmB7.GNqVo.yxtKT.allowTextSelection",
    minimizedThreadBody: "._nzWz",
    threadMessageDate: ".AL_OM.l8Tnu.I1wdR",
    emailThreadParent: ".Mq3cC",
    composeTabsParent: ".nuum5",
    selectedComposeTab: ".WkjaK.KaP9I.GlX5B",
    composeTab: ".WkjaK.GlX5B",
    dynamicConversationBody: ".L72vd", // updates with new convo: used for detecting root removal
    timeIndicator: ".AL_OM.l8Tnu.I1wdR",
    emailHeader: ".lDdSm.l8Tnu",
    iconButton: ".ms-OverflowSet-item",
  };

  parseThreadFromDom() {
    const emailObjectList = [];
    const subjectNode = document.querySelector(this.selectors.threadSubject);
    const emailMessages = document.querySelectorAll(
      this.selectors.threadMessage
    );

    if (emailMessages.length === 0) return null;

    let subject = "";
    if (subjectNode) subject = subjectNode.innerText;

    for (let emailNode of emailMessages) {
      const emailObject = {
        subject: "",
        from: { name: "", address: "" },
        to: [],
        cc: [],
        bcc: [],
        contentHtml: "",
        id:
          Math.random().toString(36).substring(2, 15) +
          Math.random().toString(36).substring(2, 15),
        timestamp: 0,
        sentAt: 0,
      };

      if (subject) {
        emailObject.subject = subject;
      }

      const sender = emailNode.querySelector(this.selectors.threadSender);
      if (sender) {
        // use the handle name address string function provided by the parent class
        emailObject.from = this.handleNameAddressString(sender.innerText);
      }

      const recipientBoxes = emailNode.querySelectorAll(
        this.selectors.recipientContainer
      );
      for (let box of recipientBoxes) {
        const boxType = box.querySelector("label").innerText;
        if (boxType === "To:") {
          const recipients = box.querySelectorAll(this.selectors.recipient);
          if (recipients.length > 0) {
            const recipientList = [];
            for (let name of recipients) {
              if (name.innerText === "You") {
                continue;
              }
              // remove all punctuation at the ends from name.innerText both ends of string and leading/trailing whitespace
              const nameString = name.innerText
                .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "")
                .trim();
              recipientList.push({ name: nameString, address: "" });
            }
            emailObject.to = recipientList;
          }
        } else if (boxType === "Cc:") {
          const recipients = box.querySelectorAll(this.selectors.recipient);
          if (recipients.length > 0) {
            const recipientList = [];
            for (let name of recipients) {
              if (name.innerText === "You") {
                continue;
              }
              // remove all punctuation from name.innerText both ends of string and leading/trailing whitespace
              const nameString = name.innerText
                .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "")
                .trim();
              recipientList.push({ name: nameString, address: "" });
            }
            emailObject.cc = recipientList;
          }
        }
      }

      const body = emailNode.querySelector(this.selectors.threadBody);
      if (body) {
        emailObject.contentHtml = body.innerHTML;
      } else {
        // try to parse minimized email (only has preview text)
        const minimizedBody = emailNode.querySelector(
          this.selectors.minimizedThreadBody
        );
        if (minimizedBody) {
          emailObject.contentHtml = minimizedBody.innerHTML;
        }
      }

      const date = emailNode.querySelector(this.selectors.threadMessageDate);
      if (date) {
        emailObject.timestamp = this.convertDateToUnix(date.innerText);
      }

      emailObjectList.push(emailObject);
    }

    const threadId = this.get.threadId();

    const threadObject = {
      emails: emailObjectList,
      thread_id: threadId,
    };

    this.cache.processedConversationMap.set(threadId, threadObject);

    return threadObject;
  }
}

class OutlookElement {
  constructor(element) {
    this.$el = $(element);
  }
}

class Compose extends OutlookElement {
  constructor(element) {
    super(element);
    this.parent = $(element).parent();
  }

  id() {
    // if it has an id, return it, else return 0
    if (this.parent.attr("id")) {
      let id = this.parent.attr("id");
      return parseInt(id.split("_")[2]);
    } else return 0;
  }

  type() {
    if (this.$el.find(".f1MMn").length > 0) {
      return "compose";
    } else {
      return "reply";
    }
  }

  find(selector) {
    return this.$el.find(selector);
  }
}
