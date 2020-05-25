
bzhao = new ActiveXObject("BZWhll.WhllObj");
bzhao.Connect( "" );
var ScreenNav = bzhao.PSGetText(9, 120);

function checkScreen() {
//checks screen to see if text on second line says main menu, then navigates
//to it. 
if (ScreenNav == "Main Menu") {
	bzhao.SendKey( "Historical<Enter>" );
	bzhao.WaitReady(10, 1);
	bzhao.SendKey( "Reports<Enter>" );
	bzhao.SendKey( "<Enter>" );
	bzhao.WaitReady(10, 1000);
	};
};

function returnMainMenu() {
	bzhao.SendKey( "<Escape>" );
	bzhao.WaitReady(10, 300);
	bzhao.SendKey( "<Escape>" );
	bzhao.WaitReady(10, 300);
	bzhao.SendKey( "<Escape>" );
	bzhao.WaitReady(10, 300);
	bzhao.SendKey( "<Escape>" );
	bzhao.WaitReady(10, 300);
};

function clearAlarms() {
	//fetch text from positon 77, 3 chars long
	var alarmCount = bzhao.PSGetText(3, 77);
	var checkScreen = bzhao.PSGetText(20, 115);
	if (checkScreen == "Alarm History Report"){
	//enters keystrokes "s, y and enter" until index iterates to alarmCount value
		if (alarmCount == 0){
			MsgBox("No Alarms to clear.");
			bzhao.WaitReady(10, 300);
			returnMainMenu();
		} else {	
			for (i = 0; i < alarmCount; i++){
				bzhao.WaitReady(10, 300);
				bzhao.SendKey( "s" );
				bzhao.WaitReady(10, 1);
				bzhao.SendKey( "y" );
 				bzhao.WaitReady(10, 1);
				bzhao.SendKey( "<Enter>" );
				bzhao.WaitReady(10, 1);
				}	
				bzhao.WaitReady(10, 300);
				MsgBox(alarmCount + " Alarms Resolved.");
				bzhao.WaitReady(10, 300);
				returnMainMenu();
		};
		} else {
	 	MsgBox("Couldn't reach Alarm History Report\nPlease return to Main Menu and run script again.");
	};
};

checkScreen();
clearAlarms();

bzhao.Disconnect();
