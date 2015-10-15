

Function  RandSelectWebList(myWebList)
    itemCount = myWebList.GetROProperty("items count")
    myIndex = randomNumber(0, itemCount - 1) 
        myWebList.Select "#" & myIndex
End Function
set myList = Browser("IE").Page("Dashboard").WebList("bloodtype")
 RandSelectWebList(myList)                                    


Function  RandSelectWebList(myWebList)
    itemCount = myWebList.GetROProperty("items count")   
selectItemNames = split(myWebList.GetROProperty("all items"), ";")
result = selectItemNames(RandomNumber(0,itemcount-1))
myWebList.Select result                                
End Function
set myList = Browser("IE").Page("Dashboard").WebList("bloodtype")
RandSelectWebList(myList)