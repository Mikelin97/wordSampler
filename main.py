import pandas as pd
import os


def check(entry, range):
    """

    :param entry: 用户的输入，必须是数字
    :param range: 用户输入数字的范围
    :return: boolean，如果用户的数字符合规定返回True，否则返回False
    """
    if not entry.isdigit():
        print('Entry must be a number ^^')
        return False
    if int(entry) < int(min(range)) or int(entry) > int(max(range)):
        print('Please enter number within range...')
        return False

    return True

def read(vocab_dir, files):
    """
    从文件夹里面读取并返回所需要的文件。
    header都设置成了1，是因为读取的单词文件有固定格式，
    把文件的第二行作为表头来读取了。
    :param vocab_dir: 文件夹的path
    :param files: list, 所需要读取的所有文件名
    :return: 一个pandas dataframe，包含了所有文件里的内容。
    """
    dataframes = []
    for file in files:
        filepath = vocab_dir + file
        df = pd.read_excel(filepath, header=1)
        dataframes.append(df)
    result = pd.concat(dataframes)
    return result

def select(dataframe, num_desired_words):
    """
    利用pandas自带的sample方程，从所有单词中随机选取。
    :param dataframe: 单词所在的dataframe
    :param num_desired_words: 希望生成单词的总数
    :return: 一个pandas dataframe,随机选取后的单词
    """
    result = dataframe.sample(int(num_desired_words))
    return result







def generate():
    """
    主要的生成方程。还属于未完工状态LOL
    :return:
    """
    current_working_directory = os.getcwd()
    vocab_directory = current_working_directory + '/data/SSAT英语单词/'
    vocab_lists = [filename for filename in os.listdir(vocab_directory) if 'Word List' in filename]
    list_category = [num.split('.')[0].split(' ')[2] for num in vocab_lists]
    list_nums = [int(value) for value in list_category if value.isdigit()]
    word_list_range = str(min(list_nums)) + '-' + str(max(list_nums))
    while True:
        mode = input('Please enter your mode selection (c for continuous, d for discrete): ')
        if mode[0] == 'c' or mode[0] == 'C':
            print('Selected Mode is: Continuous')
            mode = 0
            break
        elif mode[0] == 'd' or mode[0] == 'D':
            print('Selected Mode is: Discrete')
            mode = 1
            break
        else:
            print('Please re-enter the mode to start ><...')
            continue
    if mode == 0:
        actual_start = False
        actual_end = False
        while not (actual_start and actual_end):
            if actual_start:
                pass
            else:
                start = input('Please enter the start list number (range is: {0}) '.format(word_list_range))
                if check(start, list_nums):
                    actual_start = int(start)
                else:
                    continue
            end = input('Please enter the end list number (range is: {0}) '.format(word_list_range))
            if check(end, list_nums):
                actual_end = int(end)
            else:
                continue
            if int(actual_end) <= int(actual_start):
                print('the end must be larger than the start...')
                actual_start = False
                actual_end = False
        desired_list = []
        for i in range(actual_start, actual_end+ 1):
            for file in vocab_lists:
                if file.split('.')[0].split(' ')[2] == str(i):
                    desired_list.append(file)
        words = read(vocab_directory, desired_list)
        total_words = words.shape[0]
        while True:
            num_desired_words = input('Please enter how many words you would like to get? ')
            if check(num_desired_words, [50, total_words]):
                break
            else:
                continue
        result = select(words, num_desired_words)
        result_filename =current_working_directory + '/result/' + str(actual_start) + '-' + str(actual_end) + '随机抽查.xlsx'
        result.to_excel(result_filename, index=False)
        return


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    generate()



