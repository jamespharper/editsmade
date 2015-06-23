string1 = "Hello. Me name is James Harper. This is my text parsing and comparison program, editsmade."
string2 = "Hello. My name is James Harper. This is my text parsing and comparison program, editsmade."

words1 = string1.split()
words2 = string2.split()

diff1 = []
diff2 = []

loc = []

for i in range(len(words1)):
    if words1[i] != words2[i]:
        diff1.append(words1[i])
        diff2.append(words2[i])
        loc.append(i)
        print words1[i], "--->", words2[i]

for i in range(len(loc)-1):
    if loc[i] == loc[i + 1] - 1:
        diff1[i] += " " + diff1[i + 1]
        diff2[i] += " " + diff2[i + 1]
        del diff1[i + 1]
        del diff2[i + 1]

print diff1, diff2
