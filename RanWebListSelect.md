1,已经验证这个方法可行，以前的错误是忘记set myList = Browser......set3个字母忘记了。
Function  RandSelectWebList(myWebList)
    itemCount = myWebList.GetROProperty("items count")
    myIndex = randomNumber(0, itemCount - 1) 
        myWebList.Select "#" & myIndex
End Function
set myList = Browser("51testing软件测试网").Page("个人资料 - 51Testing软件测试论坛").WebList("bloodtype")
 RandSelectWebList(myList)                                    ’没想到函数也能这么用。

2.已经验证这个方法可行，
Function  RandSelectWebList(myWebList)
    itemCount = myWebList.GetROProperty("items count")   
selectItemNames = split(myWebList.GetROProperty("all items"), ";")
result = selectItemNames(RandomNumber(0,itemcount-1))
myWebList.Select result                                 ‘result不能加引号，为什么
End Function
set myList = Browser("51testing软件测试网").Page("个人资料 - 51Testing软件测试论坛").WebList("bloodtype")
RandSelectWebList(myList)