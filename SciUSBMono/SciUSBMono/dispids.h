// dispids.h

#pragma once

// properties
#define				DISPID_position						0x0000
#define				DISPID_currentGrating				0x0100
#define				DISPID_AmOpen						0x0101
#define				DISPID_MonoInfo						0x0102
#define				DISPID_NumberOfGratings				0x0103
#define				DISPID_GratingInfo					0x0104
#define				DISPID_ConfigFile					0x0105
#define				DISPID_autoGrating					0x0106
#define				DISPID_INIFile						0x0109
#define				DISPID_moveTime						0x010a
#define				DISPID_AmBusy						0x010b
#define				DISPID_DisplayName					0x010c
#define				DISPID_WaitForComplete				0x010d
#define				DISPID_RunCurrent					0x010e
#define				DISPID_IdleCurrent					0x010f
#define				DISPID_GratingZeroOffset			0x0111
#define				DISPID_GratingStepsPerNM			0x0112
#define				DISPID_ApplyBacklashCorrection		0x0113
#define				DISPID_MinimumPosition				0x0114
#define				DISPID_MaximumPosition				0x0115
#define				DISPID_Asynchronous					0x0116
#define				DISPID_HighSpeed					0x0117
#define				DISPID_AllowChangeZeroOffset		0x0118
#define				DISPID_SetupWindow					0x0119
#define				DISPID_RapidScanRunning				0x011a
#define				DISPID_PageSelected					0x011b
#define				DISPID_BacklashSteps				0x011c

// message handlers
#define				DISPID_OnError						0x0200
#define				DISPID_OnStatusMessage				0x0201
#define				DISPID_OnHaveNewPosition			0x0202
#define				DISPID_OnChangedGrating				0x0203
#define				DISPID_OnAmInitialized				0x0204

// methods
#define				DISPID_Setup						0x0120
#define				DISPID_SetGratingParams				0x0121
#define				DISPID_IsValidPosition				0x0122
#define				DISPID_ConvertStepsToNM				0x0123
#define				DISPID_Abort						0x0124
#define				DISPID_WriteConfig					0x0125
#define				DISPID_WriteINI						0x0126
#define				DISPID_GetCounter					0x0127
#define				DISPID_setInitialPositions			0x0128
#define				DISPID_SaveGratingZeroPosition		0x0129
#define				DISPID_doInit						0x012b
#define				DISPID_CanAutoSelect				0x012c
#define				DISPID_SetMotorControl				0x012d
#define				DISPID_ConvertSteps					0x012e
#define				DISPID_ConvertNM					0x012f
#define				DISPID_DisplayConfigValues			0x0130
#define				DISPID_StartRapidScan				0x0131
#define				DISPID_RemoveNamedObject			0x0132

// events
#define				DISPID_RequestMainWindow			0x0140
#define				DISPID_Error						0x0141
#define				DISPID_StatusMessage				0x0142
#define				DISPID_HaveNewPosition				0x0143
#define				DISPID_ChangedGrating				0x0144
#define				DISPID_BusyStatusChange				0x0145
#define				DISPID_RequestChangeGrating			0x0146
#define				DISPID_QueryAllowChangePosition		0x0147
#define				DISPID_QueryAllowChangeZeroOffset	0x0148
#define				DISPID_RapidScanStepped				0x0149
#define				DISPID_RapidScanEnded				0x014a

// host object
#define				DISPID_Host_SendCommand				0x0180
#define				DISPID_Host_doDelay					0x0181