dim message,speak
message=inputbox("Enter Text Here","Speaking Software")
set speak=createobject("sapi.spvoice")
speak.speak message