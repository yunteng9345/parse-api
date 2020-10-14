# 判断是否有下划线, 有的话转为驼峰格式
# def isunderline(text):
#         for char in text:
#             if char == "_":
#             print(char)
#
#     pass


def to_camel_case(snake_str):
    if "_" in text:
        components = snake_str.split('_')
        # We capitalize the first letter of each component except the first one
        # with the 'title' method and join them together.
        return components[0] + ''.join(x.title() for x in components[1:])


text = "uepay_order_request"
print(to_camel_case(text))
