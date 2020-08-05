var='n0'
def fun():
    global var
    var = "hello"
    var2= "lol"
    return var2
print(var)
fun()
print(var)
print(fun())
