n=int(input())
a=list(map(int,input().split()))
b=list(map(int,input().split()))
countryset=(list(set(a)))
countryset.sort()
d,d1={},{}
for i in countryset:
    d[i]=0
    d1[i]=0
for i in range(n):
    d[a[i]]+=b[i]
for i in a:
    d1[i]+=1
for i in range(len(countryset)):
    d[countryset[i]]=d[countryset[i]]//d1[countryset[i]]
for i in range(n):
    b[i]=b[i]-d[a[i]]
print(*b,sep=' ')