valid = False
while not valid:
    try:
        txt = input("enter text: ")
    except:
        print ("entry error")
    else:
        if txt.count(" ") == len(txt):
            print ("please enter some text")
        else:
            print ("thank you, your text is " + txt)
            valid = True


