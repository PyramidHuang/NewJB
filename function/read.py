def read_file(file):
    """
    读取数据
    :param string file:路径名称
    :return:每一行数据的列表
    """
    data = []
    with open(file, "r", newline="") as filereader:
        for row in filereader:
            data.append(row)
    return data


