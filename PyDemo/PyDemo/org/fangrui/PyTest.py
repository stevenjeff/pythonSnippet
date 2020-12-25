'''
@author: admin test orgin test2dd test2dd
'''
from _collections import defaultdict
from string import Template

from fabric import Connection


def main():
    # ip ����������
    # �����ĵ�������ssh�������¼���Ͳ���Ҫ connect_kwargs ��ָ�������ˡ�
    c = Connection("root@232.231.231.22", connect_kwargs={"password": "youpassword"})
    with c.cd('/var/www/youproject'):
        c.run("git pull origin master")
        c.run("/usr/bin/supervisorctl -c ../supervisor/supervisord.conf restart youproject")


def pythonic():
    #  列表推导式
    chars = [c for c in 'python']
    print(chars)
    # 字典推导式
    dicts = {'a': 1, 'b': 2, 'c': 3, 'd': 4, 'e': 5}
    newDict = {k: v * 3 for (k, v) in dicts.items()}
    print(newDict)
    # 集合推导式
    sets = {1, 2, 3, 4, 5, 6, 7}
    new_sets = {i * 4 for i in sets}
    print(new_sets)
    # 合并字典
    dict_a = {'a': 1, 'b': 3}
    dict_b = {'c': 2, 'd': 4}
    dict_all = {**dict_a, **dict_b}
    print(dict_all)
    # 复制列表
    nums = [1, 2, 3, 5]
    new_nums = nums[::]
    nums2 = nums
    nums = [3, 4, 5, 6]
    print(new_nums)
    print(nums2)
    print(nums)
    # 反转列表
    reverse_nums = [1, 2, 3, 4, 5, 6]
    reverse_nums = reverse_nums[::-1]
    print(reverse_nums)
    # 变量交换
    a, b = 1, 2
    a, b = b, a
    print(f'{a}:{b}')
    # 高级拆包
    a, *b, c = 1, 2, 3, 4, 5, 6
    print(b)

    # 函数返回多个值（其实是自动packing成元组）然后unpacking赋值给4个变量
    def f():
        return 1, 2, 3, 4, 5

    a, b, c, d, e = f()
    print(f'{a} ,{b} ,{c} ,{d}, {e}')
    # 列表合并成字符串
    temp = " ".join(["I", "Love", "Python"]);
    print(temp)
    # 链式比较
    if a > 2 and a < 5:
        pass
    if 2 < a < 5:
        pass

    # 字典代替多个if else
    def fun(x):
        if x == 'a':
            return 1
        elif x == 'b':
            return 2
        else:
            return None

    def funDict(x):
        return {"a": 1, "b": 2}.get(x)

    print(fun('d'))
    # 有下标索引的枚举
    for i, e in enumerate(["a", "b", "c"]):
        print(i, e)
    # 生成器
    g = (i ** 2 for i in range(5))
    print(g)
    for i in g:
        print(i)
    # 默认字典 defaultdict
    d = dict()
    # d['nums'] error
    d = defaultdict(list)
    print(d["nums"])
    # 列表中出现次数最多的元素
    nums = [1, 2, 3, 3]
    maxcount = max(set(nums), key=nums.count)
    print(maxcount)
    # 读写文件
    with open("text.txt", "w") as f:
        f.writelines("hello")
    # 判断对象类型，可指定多个类型
    print(isinstance(maxcount, (int, str)))
    # 类似的还有字符串的 startswith，endswith
    temp = "http://foolish.net".startswith(("http", "https"))
    print(temp)

    # 使用装饰器
    def makebold(f):
        return lambda: '<b>' + f() + '</b>'

    def makeitalic(f):
        return lambda: '<i>' + f() + '</b>'

    @makebold
    @makeitalic
    def sayHello():
        return "hello world"

    print(sayHello())

    def equalTest():
        f = Foo();
        print(f == None)
        print(f is None)

    equalTest()

    def stringOperate():
        print('Hello' + ' ' + 'World' + '!')
        strlist = ['Hello', ' ', 'World', '!']
        print(''.join(strlist))
        print('{} {}!'.format('Hello', 'World'))
        print('%s %s!' % ('Hello', 'World'))
        str = (
            'Hello'
            ' '
            'World'
            '!'
        )
        print(str)
        s = Template('${s1} ${s2} !')
        print(s.safe_substitute(s1='Hello', s2='World'))
        s1 = 'Hello'
        s2 = 'World'
        print(f'{s1} {s2}   !')

        def power(x):
            return x * x

        x = 6
        print(f'{x} * {x} = {power(x)}')

    stringOperate()


class Foo(object):
    def __eq__(self, other):
        return True


if __name__ == '__main__':
    pythonic()
