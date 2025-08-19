Private Sub S2KVBAScriptingObject_OnControl()
If (BstrOutputPlug_1.HasControled = True) Then
 
	if BoolParam1 Then
		BstrInputPlug_1.ControlValue = ProcessInNibblesB2L(BstrOutputPlug_1.NewValue)
	Else
		BstrInputPlug_1.ControlValue = BstrOutputPlug_1.NewValue
	End If
	BstrInputPlug_1.ControlQuality = S2K_QUALITY_GOOD
	BstrInputPlug_1.ControlTimeStamp = BstrOutputPlug_1.NewTimeStamp
End If
 
End Sub

------------------

Private Sub S2KVBAScriptingObject_OnChange()
        If (BstrInputPlug_5.HasChanged = True) Then
			if BoolParam1 Then
				BstrOutputPlug_5.NewValue = ProcessInNibblesL2B(BstrInputPlug_5.Value)
			Else
				BstrOutputPlug_5.NewValue = BstrInputPlug_5.Value
			End If
			BstrOutputPlug_5.NewQuality = S2K_QUALITY_GOOD
			BstrOutputPlug_5.NewTimeStamp = Now
        End If
End Sub