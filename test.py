

file_path = input()
print(file_path)
print("path?")
with open(file_path, 'r') as file :
    data = file.read()
print(data)
