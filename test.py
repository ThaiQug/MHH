# Import thư viện cần dùng (numpy, openpyxl)
import numpy as np
import openpyxl
from gamspy import Container, Set, Parameter, Variable, Equation, Model, Sum, Sense

# Initialize các giá trị ban đầu
n = 8 # Số loại sản phẩm
m = 5 # Số loại nguyên liệu

# lấy dữ liệu từ file xlsx sử dụng hàm load_workbook của thư viện openpyxl
wb = openpyxl.load_workbook('ABCDE.xlsx')
sheet = wb['Sheet1']

# Lấy các giá trị trong sheet
L = sheet['B2:I2'] # giá sản xuất sản phẩm
Q = sheet['B3:I3'] # giá bán sản phẩm
B = sheet['B7:F7'] # giá mua nguyên liệu
S = sheet['B8:F8'] # giá bán lại nguyên liệu
A = sheet['B12:F19'] # số nguyên liệu cần dùng để sản xuất sản phẩm

container=Container() # Tạo một bao đóng để quản lý sử dụng class Container

# Chuyển các giá trị vừa lấy được ở trên và chuyển chúng thành các mảng (array)
l = np.array([float(cell.value) for cell in L[0]])
q = np.array([float(cell.value) for cell in Q[0]])
b = np.array([float(cell.value) for cell in B[0]])
s = np.array([float(cell.value) for cell in S[0]])
a = np.array([[float(cell.value) for cell in row] for row in A])

# Tạo nhu cầu d
d = np.random.binomial(10, 1/2, size=(2, 8))
p = np.array([0.5, 0.5])

i = Set(container, "product", records=[ v for v in range(n) ])
j = Set(container, "preordered_part", records=[ v for v in range(m) ])
k = Set(container, "scenarios", records=["1", "2"])

# Nhập tham số
va = Parameter(container, "va", domain=[i, j], records=a)
vq = Parameter(container, "vq", domain=[ i ], records=q)
vl = Parameter(container, "vl", domain=[ i ], records=l)
vb = Parameter(container, "vb", domain=[ j ], records=b)
vs = Parameter(container, "vs", domain=[ j ], records=s)
vd = Parameter(container, "vd", domain=[k, i], records=d)
vp = Parameter(container, "vp", domain=[k], records=p)
 
# Tạo biến
x = Variable(container, "x", domain=[j], type="integer")
y = Variable(container, "y", domain=[k, j], type="integer")
z = Variable(container, "z", domain=[k, i], type="integer")

# In tham số
print("va: ")
print( va.records )
print("--------------------------")
print("vq: ")
print( vq.records )
print("--------------------------")
print("vl: ")
print( vl.records )
print("--------------------------")
print("vb: ")
print( vb.records )
print("--------------------------")
print("vs: " )
print( vs.records )
print("--------------------------")
print("vd: ")
print( vd.records )
print("--------------------------")
print("vp: ")
print( vp.records )

# Tạo ràng buộc
st1 = Equation(container, "st1", domain=[k, j])
st2 = Equation(container, "st2", domain=[k, i])
st21 = Equation(container, "st21", domain=[k, i])
st3 = Equation(container, "st3", domain=[k, j])
st4 = Equation(container, "st4", domain=[j])
st1[k, j] = y[k, j] == x[j] - Sum(i, va[i, j] * z[k, i])
st2[k, j] = 0 <= z[k, i] <= vd[k, i]
st21[k,i] = z[k, i] <= vd[k, i]
st3[k ,j] = y[k, j] >=0
st4[j] = x[j] >= 0

 # Tạo phương trình mục tiêu
objective_fucntion = Sum(j, vb[j] * x[j]) + Sum(k, vp[k] * (Sum(i, (vl[i] - vq[i]) * z[k, i]) - Sum(j, vs[j] * y[k, j])))

question = Model(container, name="question1", equations=[st1,st2], problem="MIP", sense=Sense.MIN, objective=objective_fucntion)

question.solve()

print( question.objective_value )