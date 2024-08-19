import os


def of_xml(magazin, date, time, count_in, count_out):
    head = '<?xml version="1.0"?>\n<CM_StorageEnters>'
    footer = '</CM_StorageEnters>'
    date_of_name = f'{date[6:]}{date[3:5]}{date[0:2]}'
    name = f'{magazin}-{date_of_name}-1600.xml'
    with open(f'\\\\192.168.10.55\\Counters\\{magazin}\\{name}', "w") as file:
        file.write(head)
        for i in range(len(time)):
            file.write(
                f'<StorageEnter>\n<ID_Enter>{magazin}</ID_Enter>\n<TimeRecord>{date}-{time[i]}</TimeRecord>\n<_Interval>600</_Interval>\n<SumIn>{count_in[i]}</SumIn>\n<SumOut>{count_out[i]}</SumOut>\n</StorageEnter>'
            )

        file.write(footer)
    return '200'

def delete_files_in_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f'Ошибка при удалении файла {file_path}. {e}')
    return '300'

