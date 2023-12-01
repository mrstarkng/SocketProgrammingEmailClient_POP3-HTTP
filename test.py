

file_path = input()
print(file_path)
print("C:\\Users\\ASUS ZenBook\\OneDrive\\Máy tính\\AVL.cpp")
with open(file_path, 'r') as file :
    data = file.read()
print(data)