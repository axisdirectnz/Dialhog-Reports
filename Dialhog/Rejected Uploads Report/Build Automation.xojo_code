#tag BuildAutomation
			Begin BuildStepList Linux
				Begin BuildProjectStep Build
				End
			End
			Begin BuildStepList Mac OS X
				Begin BuildProjectStep Build
				End
				Begin SignProjectStep Sign
				  DeveloperID=
				End
			End
			Begin BuildStepList Windows
				Begin IDEScriptBuildStep SaveScript , AppliesTo = 0, Architecture = 0, Target = 0
					doCommand("SaveFile")
				End
				Begin BuildProjectStep Build
				End
			End
#tag EndBuildAutomation
