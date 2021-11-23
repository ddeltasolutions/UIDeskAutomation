# UIDeskAutomation
This is a .NET library that can be used to automate Windows desktop programs based on their user interface. It is created on top of managed Microsoft UI Automation API.

Here is an example. You can start an application like this:

<i><b>using UIDeskAutomationLib;<br>
...<br>
var engine = new Engine();<br>
engine.StartProcess("notepad.exe");<br></b></i>

After starting the application, you can set text to Notepad like this:

<i><b>UIDA_Window notepadWindow = engine.GetTopLevel("Untitled - Notepad");<br>
UIDA_Document document = notepadWindow.Document();<br>
document.SetText("Some text");<br></b></i>

You can find more information about this library <a href="http://automationspy.freecluster.eu/uideskautomation.html">HERE</a>.<br>
Install the library from <a href="https://www.nuget.org/packages/UIDeskAutomation/">Nuget</a>.
