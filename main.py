from datetime import datetime
print(datetime.today().strftime('%Y-%m-%d'))
print(type(datetime.today().strftime('%Y-%m-%d')))
class ahmed:
        pass
say = ahmed()
print(type(say))


class dog:
    #class object attribute
    species = 'mammal'

    def __int__(self, breed,name):
        self.breed = breed
        self.name = name


sam = dog()
sam.breed = 'huskie'
print(sam.species)
frank = dog()
print(sam.breed)

class circle:
    pi = 3.14

    def __init__(self, radius = 1):
        self.radius = radius
        self.area = radius * radius * circle.pi
    def Set_radius(self, new_radius):
        self.radius = new_radius
        self.area = new_radius * new_radius * circle.pi
    def circum(self):
        circ = 2 * self.radius *self.pi
        return circ

circle1 = circle();

print(circle1.circum())

circle1.Set_radius(1)

print(circle1.circum())

print(circle1.area)

class Ahmed:
    def __init__(self):
        print("Ahmed ELsayed")
    def person(self):
        print("Hero")

ahmed = Ahmed()
print(ahmed.__init__())
print(ahmed.person())

def erlang_b(b, a):
    p = 1
    for i in range(0, b):
        p *= (a / (a + i))
    return (b / (b + a)) * p

b = float(input("Enter the number of servers (b): "))
a = float(input("Enter the offered load (a): "))
result = erlang_b(int(b), a)

print("The probability of blocking is: ",result)