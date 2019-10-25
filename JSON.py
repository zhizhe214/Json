import json
import 用Python完成Excel的常用操作

originJson='{"error_code":17,"stu_info":[{"id":309,"name":"小白","sex":"男","age":28,\
"addr":"河南省济源市北海大道32号","grade":"天蝎座","phone":"18512572946","gold":100},\
{"id":310,"name":"小白","sex":"男","age":28,"addr":"河南省济源市北海大道32号",\
"grade":"天蝎座","phone":"18516572946","gold":100}]}'

'''
# loads()函数是将json格式数据转换为字典
# （可以这么理解，json.loads()函数是将字符串转化为字典）
res=json.loads(originJson)
print(res)                           # 打印字典
print(type(res))                     # 打印 res 的类型
print(res.keys())                    # 打印字典的所有 key
print(res["error_code"])             # 访问字典里的值
print(res["stu_info"])               # 访问字典里的值
'''

 # json.dumps() 函数，将字典转化为 json 字符串
dict_1={'StudentNum':'24','Name':'Peak','age':'17'}
json_string=json.dumps(dict_1)
print(json_string)
print('dict_1的类型为： '+str(type(dict_1)))
print('json_string的类型为： '+str(type(json_string)))

pi=3.14159265359
print('(%8.2f)'%(pi))