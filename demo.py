
def get_whole_book():
    file_name= input("What file should I open")
    file= open(file_name,'r')
    all_lines= file.readlines()
    return all_lines

def count_words(all_lines):
    word_counter ={}
    punctuation ='"`~!@#$%^&*()-_=+[]{}|:;,.<>/?\n\r'
    for line in all_lines:
        words_in_line= line.split(" ")
        for word in words_in_line:
            word= word.strip(punctuation,'')
            word = word.lower()
            if word in word_counter:
                word_counter[word] = word_counter[word]+1
            else:
                word_counter[word]= 1

    report_results(word_counter)

def report_results(word_counts):
    for word in word_counts:
        print(f"{word} : {word_counts[word]}")





def main():
    book_lines= get_whole_book()
    count_words(book_lines)

main()



