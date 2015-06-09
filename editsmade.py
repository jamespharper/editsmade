string1 = "Hello. My name is James Harper. This is my text parsing and comparison program, editsmade."
string2 = "Hello. My name is James Harper. This is my text parsing and comparison program, editsmade."

words = []
start = 0

for i in range(len(string1)):
    if string1[i] == " ":
        if start == 0:
            start = i
        elif start != 0:
            print i
            print string1[start:i + 1]
            words.append(string1[start:(i + 1)])
            start = 0
        else:
            print "ERROR"
