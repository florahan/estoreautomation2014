'1,�Ѿ���֤����������У���ǰ�Ĵ���������set myList = Browser......set3����ĸ�����ˡ�

Function  RandSelectWebList(myWebList)
    itemCount = myWebList.GetROProperty("items count")
    myIndex = randomNumber(0, itemCount - 1) 
        myWebList.Select "#" & myIndex
End Function
set myList = Browser("IE").Page("Dashboard").WebList("bloodtype")
 RandSelectWebList(myList)                                    ��û�뵽����Ҳ����ô�á�

'2.�Ѿ���֤����������У�
Function  RandSelectWebList(myWebList)
    itemCount = myWebList.GetROProperty("items count")   
selectItemNames = split(myWebList.GetROProperty("all items"), ";")
result = selectItemNames(RandomNumber(0,itemcount-1))
myWebList.Select result                                 'result���ܼ����ţ�Ϊʲô
End Function
set myList = Browser("IE").Page("Dashboard").WebList("bloodtype")
RandSelectWebList(myList)