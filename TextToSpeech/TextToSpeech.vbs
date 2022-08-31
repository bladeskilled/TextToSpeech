dim message,speak
message=inputbox("Enter Text Here","Text To Speech")
set speak=createobject("sapi.spvoice")
speak.speak message
