# anyCAN

```
 ████ ████ ████ ████ ████ ████ ████ ████ ████╗
╔════════════════════════════════════════════╝
║
║    █████╗ ███╗   ██╗██╗   ██╗
║   ██╔══██╗████╗  ██║██║   ██║
║   ███████║██╔██╗ ██║██║██ ██║
║   ██╔══██║██║╚██╗██║  ╚═══██║                     
║   ██║  ██║██║ ╚████║ ██████╔╝
║   ╚═╝  ╚═╝╚═╝  ╚═══╝ ╚═════╝ 
║
║       █████╗ █████╗ ███╗   ██╗
║     ██╔════╝██╔══██╗████╗  ██║
║     ██║     ███████║██╔██╗ ██║
║     ██║     ██╔══██║██║╚██╗██║
║     ╚██████╗██║  ██║██║ ╚████║  █ ╗  █ ╗  █ ╗        { .py script to automate testing TestCases(written in excel) }
║      ╚═════╝╚═╝  ╚═╝╚═╝  ╚═══╝  ╚ ╝  ╚ ╝  ╚ ╝        
╚ ════ ════ ════ ════ ════ ════ ════ ════ ════ ═╝
```

anyCAN_Tx   →   For Single Test Case Mode (Single .xlsx)  
anyCAN      →   For Multi  Test Case Mode (Folder of one or more .xlsx)     


We have a few keyboard shortcuts within the script(s), see below:    
**<< NOTE: The shortcuts are common for all the Tx scripts >>**


**SHORTCUTS**:

**ESC**          :   Pause/Resume the CAN log being displayed in the terminal.   
**Alt + S**      :   Open Tx GUI   
**Ctrl + P**     :   Pause/Resume the current Tx loop (or) TestCase                                                             



**Tx GUI window:**

![image](https://github.com/user-attachments/assets/fb5069b3-7a40-4dc5-9c3e-9e956550231d)     

**MODES:**   
**<< NOTE: The modes only are an option in the "anyCAN" script & not present in "anyCAN_Tx" >>**
  
**1) AUTO:** After "TestCase folder" is uploaded,  1st TestCase is automatically loaded in GUI and "Send All" needs to be pressed by the user to start Tx... Once 1st TestCase is finished, next TestCase is automatically loaded and executed after 5 seconds. This repeats until all TestCases are completed.

**2) MANUAL:** Only difference from auto mode is that, after a TestCase is done, the next TestCase will automatically load into the GUI but we need to press "send all" manually each time in order to start execution.


