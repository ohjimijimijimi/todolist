var COL_TASK = 1;
var COL_PRIORITY = 2;
var COL_EXPIRE = 3;
var COL_USER = 4;
var COL_STATUS = 5;
var COL_PROGRESS = 6;

var SHEET_TASK = "Task";
var SHEET_USERS = "Users";
var SHEET_NOTS = "Notifications";
var SHEET_TOKEN = "Tokens";
var SHEET_NOTS_QUEUE = "Notification Queue";
var SHEET_LOG = "Logs";

var LOG_NOTICE = 'notice';
var LOG_WARNING = 'warning';
var LOG_ERROR = 'error';

var tokens = [{token:'%team', type:'static', position:'B2'}, 
              {token:'%recipients', type:'dynamic'},
              {token:'%link', type:'code', value:SpreadsheetApp.getActiveSpreadsheet().getUrl()},
              {token:'%user', type:'code', value:Session.getUser().getUserLoginId()},
              {token:'%task', type:'dynamic'},
             ];

/* Events start */
function onOpen() {
  Logger.log("entering OnOpen");
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var menuEntries = [ {name: "Test Mail", functionName: "test_mail"}];
  //ss.addMenu("Test", menuEntries);
  Logger.log("leaving OnOpen");
}

function onEdit(event)
{
  //Logger.log("entering onEdit");
  if (event.source.getActiveSheet().getSheetName() != SHEET_TASK) {
    return;;
  }
  var r = event.source.getActiveRange();
  dispatchEditActions(r, event);
  //Logger.log("leaving onEdit");
}

/** 
 * Edit Event Dispatcher. 
 * Parse the possibile actions triggered by the event onEdit.
 */
function dispatchEditActions(r, event) {
  //Logger.log("entering dispatchEditActions");
  switch (r.getColumnIndex()) {
    case COL_USER:
      onEditUser(r, event);
      break;;
    case COL_PROGRESS:
      onEditProgress(r, event);
      break;;
    case COL_STATUS:
      onEditStatus(r, event);
      break;;
  }
  //Logger.log("leaving dispatchEditActions");
}

/** 
 * Action onEditUser. Enqueue a notification if a task will
 * be assigned to a user.
 */
function onEditUser(r, event) {
  Logger.log("entering onEditUser");
  var task = getTask(r.getRowIndex());
  var user = getUser(task);
  enqueueNotification(task, user, 'U0');
  Logger.log("leaving onEditUser");
}

function onEditStatus(r, event) {
  //Logger.log("entering onEditStatus");
  var task = getTask(r.getRowIndex());
  var user = getUser(task);
  enqueueNotification(task, user, 'U1');
  //Logger.log("leaving onEditStatus");
}

function onEditProgress(r, event) {
  //Logger.log("entering onEditProgress");
  var task = getTask(r.getRowIndex());
  var user = getUser(task);
  enqueueNotification(task, user, 'U1');
  //Logger.log("leaving onEditProgress");
}

/* Events end */

function getPermissions(s) {
  //Logger.log("entering getPermissions");
  var values = s.getRange(2, 3, 2, s.getLastColumn()).getValues();
  var permissions = {};
  for (var i=0; i<=s.getLastColumn()-3; i++) {
    permissions[values[1][i]] = {label: values[0][i], enabled: false};
  }
  //Logger.log("leaving getPermissions");
  return permissions;
}

function getUser(task) {
  //Logger.log("entering getUser");
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);

  var permissions = getPermissions(s);
  var values = s.getRange(4, 1, s.getLastRow(), s.getLastColumn()).getValues();
  var user = {};
  for (i in values) {
    if (task.assignee == values[i][0]) {
      user['name'] = values[i][0];
      user['mail'] = values[i][1];
      permissions.U0.enabled = values[i][2] == 'yes' ? true : false;
      permissions.U1.enabled = values[i][3] == 'yes' ? true : false;
      permissions.U2.enabled = values[i][4] == 'yes' ? true : false;
      permissions.U3.enabled = values[i][5] == 'yes' ? true : false;
      user['permissions'] = permissions;
    }
  }
  //Logger.log("leaving getUser");
  return user;
}

function enqueueNotification(task, user, type) {
  Logger.log("entering enqueueNotification");
  if (checkNotificationData(task, type)) {
  var notification = getNotification(task, user, type);
    if (checkNotificationPermission(user, type)) {
      var queue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTS_QUEUE);
      queue.insertRowAfter(queue.getLastRow());
      var r = queue.getRange(queue.getLastRow()+1, 1, 1, 4);    
      r.setValues([[notification.subject, notification.body, user.mail, notification.type]]);
    }
  }
  Logger.log("leaving enqueueNotification");
}

function getNotification(task, user, type) {
  Logger.log("entering getNotification");
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTS);
  var values = s.getRange(2, 1, s.getLastRow(), s.getLastColumn()).getValues();
  var notification = {};
  for (i in values) {
    if (values[i][0] == type) {
      notification['type'] = type;
      notification['subject'] = values[i][1];
      notification['body'] = values[i][2];
    }
  }
  notification = notificationTemplate(notification, task, user);
  Logger.log("leaving getNotification");
  return notification;
}

function notificationTemplate(notification, task, user) {
  //Logger.log("entering notificationTemplate");
  for (i in tokens) {
    var token = tokens[i];
    switch (token.type) {
      case 'dynamic':
        notification = notificationTemplateDynamic(notification, token, task, user);
        break;;
      case 'code':
        notification = notificationTemplateCode(notification, token);
        break;;
      case 'static':
        notification = notificationTemplateStatic(notification, token);
        break;;
    }
    //Logger.log("leaving notificationTemplate");
  }
  return notification;
}

function notificationTemplateDynamic(notification, token, task, user) {
  //Logger.log("entering notificationTemplateDynamic");
  var re = new RegExp(token.token, 'g');
  switch (token.token) {
    case "%recipients":
      var value = user.mail;
      notification.subject = notification.subject.replace(re, value);
      notification.body = notification.body.replace(re, value);  
      break;;
    case "%task":
      var value = task.name;
      notification.subject = notification.subject.replace(re, value);
      notification.body = notification.body.replace(re, value);
      break;;
  }
  //Logger.log("leaving notificationTemplateDynamic");
  return notification;
}

function notificationTemplateCode(notification, token) {
  //Logger.log("entering notificationTemplateCode");  
  var re = new RegExp(token.token, 'g');
  notification.subject = notification.subject.replace(re, token.value);
  notification.body = notification.body.replace(re, token.value);  
  //Logger.log("leaving notificationTemplateCode");  
  return notification;
}

function notificationTemplateStatic(notification, token) {
  //Logger.log("entering notificationTemplateStatic");  
  var re = new RegExp(token.token, 'g');
  var value = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TOKEN).getRange(token.position).getValue();
  notification.subject = notification.subject.replace(re, value);
  notification.body = notification.body.replace(re, value);
  //Logger.log("leaving notificationTemplateStatic");
  return notification;
}

function getTask(row_index) {
  //Logger.log('entering getTask');
  var task_s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASK);
  var values = task_s.getRange(row_index, 1, 1, task_s.getLastColumn()).getValues();
  //Logger.log('leaving getTask');
  return {name:values[0][0], 
          expire:values[0][2], 
          assignee:values[0][3], 
          status:values[0][4], 
          progress:values[0][5]};
}

/*function checkU1Conditions(r, col) {
  //Logger.log("entering checkU1Conditions");
  switch (col) {
    case COL_PROGRESS:
      if ((r.getColumnIndex() ==  COL_PROGRESS) 
        && (r.getValue() == '100')
        && (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASK).getRange(r.getRowIndex(), COL_STATUS, 1, 1).getValue() == 'Complete')) {
        //Logger.log("leaving checkU1Conditions");
        return true;
      }
      return false;
    case COL_STATUS:
      if ((r.getColumnIndex() == COL_STATUS)
        && (r.getValue() == 'Complete')
        && (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASK).getRange(r.getRowIndex(), COL_PROGRESS, 1, 1).getValue() == '100')) {
        //Logger.log("leaving checkU1Conditions");
        return true;
      }
      return false;
  }
  //Logger.log("leaving checkU1Conditions");
  return false;
}*/

function sendNotification() {
  Logger.log("entering sendNotification");
  var queue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTS_QUEUE);
  var values = queue.getRange(2, 1, queue.getLastRow()-1, queue.getLastColumn()).getValues();
  var to_delete = [];
  for (i in values) {
    var notification = {subject: values[i][0], body: values[i][1], recipients: values[i][2], type: values[i][3]};
    try {
      MailApp.sendEmail(notification.recipients, notification.subject, notification.body);
      log("Notification sent to " + notification.recipients, LOG_NOTICE);
      to_delete.push(Number(i) + 2);
    } catch (e) {
      log(e, LOG_ERROR);
    }
  }
  // clear notification queue
  for (var i = to_delete.length-1; i>=0; i--) {
    Logger.log("deleting row " + to_delete[i]);
    queue.deleteRow(to_delete[i]);
  }
  Logger.log("leaving sendNotification");
}

/**
 * Logger function. Use it with caution.
 * This function write a log in sheet called 'Logs'.
 */
function log(msg, level) {
  if (!level) {
    level = LOG_WARNING;
  }
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  log.insertRowBefore(2);
  var range = log.getRange(2, 1, 1, 3);
  range.setValues([[new Date(), level, msg]]);
}

function checkNotificationData(task, type) {
  Logger.log('entering checkNotificationData');
  var check = eval(type+'_checkData('+Utilities.jsonStringify(task)+')');
  Logger.log('leaving checkNotificationData');
  return check;
}

function checkNotificationPermission(user, type) {
  Logger.log('entering checkNotificationPermission');
  var check = eval(type+'_checkPermission('+Utilities.jsonStringify(user)+')');
  Logger.log('entering checkNotificationPermission');
  return check;
}

function doGet() {
  var app = UiApp.createApplication();
  app.add(app.loadComponent("Add Task")).setHeight(400).setWidth(700);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

function addTaskPriorityOnClick() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRIORITY);
  return s.getRange(2, 1, s.getLastRow(), 1).getValues();
}

function addTaskButtonResetOnClick() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

/**
 * U0 Notification functions
 * Receive a mail when someone assign a task to you.
 */
function U0_checkData(task) {
  return task.name && task.status != 'Complete';
}

function U0_checkPermission(user) {
  return user.permissions.U0.enabled;
}

/** 
 * U1 Notification functions
 * Receive a mail when a task is completed.
 */
function U1_checkData(task) {
  return task.name && task.status == 'Complete' && task.progress == '100';
}

function U1_checkPermission(user) {
  return user.permissions.U1.enabled;
}

/** 
 * U2 Notification functions
 * Receive a mail when a new task is created.
 */
function U2_checkData(task) {
  return task.name && task.status == 'New' && task.progress == '0';
}

function U2_checkPermission(user) {
  return user.permissions.U2.enabled;
}

