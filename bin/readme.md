# tlbinf32.dll in a 64bits .Net application
## [install] (https://stackoverflow.com/questions/42569377/tlbinf32-dll-in-a-64bits-net-application)

1. open Windows' "Component Services"
2. open nodes to "My Computer/COM+ Applications"
3. right-click, choose to add a new Application
4. choose an "empty application", name it "tlbinf" for example
5. make sure you choose "Server application" (means it will be a surrogate that the wizard will be nice to help you create)
6. choose the user you want the server application to run as (for testing you can choose interactive user but this is an important decision to make)
7. you don't have to add any role, not any user
8. open this newly created app, right-click on "Components" and choose to add a new one
9. choose to install new component(s)
10. browse to your tlbinf32.dll location, press "Next" after the wizard has detected 3 interfaces to expose
11 That's it. You should see something like this:


Now you can use the same client code and it should work. Note the performance is not comparable however (out-of-process vs in-process).

The surrogate app you've just created has a lots of parameters you can reconfigure later on, with the same UI. You can also script or write code (C#, powershell, VBScript, etc.) to automate all the steps above.
