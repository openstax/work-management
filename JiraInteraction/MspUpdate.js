// JavaScript source code
var location = window.location.pathname;
var path = location.substring(0, location.lastIndexOf("/"));
var directoryName = path.substring(path.lastIndexOf("/") + 1);
var ExePth = "file:" + directoryName + "/JiraInteraction.exe";

MyObject = new ActiveXObject("WScript.Shell")
MyObject.Run(ExePth);
