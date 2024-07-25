import sys, os, json, asyncio, datetime, openpyxl, progressbar, keyboard
from telethon import TelegramClient
from telethon.tl.types import UserStatusOnline, UserStatusOffline, UserStatusRecently, UserStatusLastWeek, UserStatusLastMonth
from telethon import TelegramClient
from telethon.tl.functions.channels import GetParticipantsRequest
from telethon.tl.functions.contacts import ResolveUsernameRequest
from telethon.tl.types import ChannelParticipantsSearch, InputChannel

script_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
config_file_path = os.path.join(script_dir, 'config.json')
session_file_path = os.path.join(script_dir, 'session\\client.session')
json_file_path = os.path.join(script_dir, 'parsing_data\\data.json')

limit = 100

async def main():
    
    try:
        os.system('cls||clear')
        if not os.path.exists(session_file_path):
            
            print("https://my.telegram.org/apps")
            api_id = input("API ID: ")
            api_hash = input("API Hash: ")
            to_config = {'api_id': api_id, 'api_hash': api_hash, 'limit_messages': 5000}
            
            print("Save settings...")
            with open(config_file_path, 'w') as f:
                f.write(json.dumps(to_config))
            print("Save settings: Done!")
                    
            session = False
            while not session:
                print("Сreate .session file...")
                session = await create_session(api_id, api_hash)

        link = input("Link: ")
        await parsing(link)
    
    except Exception as e:
        print(f"Error: {e}\nDelete \"session\\client.session\" file and restart this program.")
        while True:
            pass

# ==================================================================================================== #

async def create_session(api_id, api_hash):
    
    async with TelegramClient(session_file_path, api_id=api_id, api_hash=api_hash) as client:
        
        if await client.is_user_authorized():
            await client.disconnect()
            os.system('cls||clear')
            print("Сreate .session file: Success!")
            return True
        else:
            await client.disconnect()
            os.system('cls||clear')
            print("Error: .session don`t valid!")
            print("Delete .session file...")
            await asyncio.sleep(5)
            os.remove(session_file_path)
            print("Delete .session file: Done")
            print("Restart registration...")
            return False

# ==================================================================================================== #

async def parsing(link):
    
    with open('config.json') as f:
        file_content = f.read()
        config = json.loads(file_content)
    
    api_id = config['api_id']
    api_hash = config['api_hash']
    
    async with TelegramClient(session_file_path, api_id=int(api_id), api_hash=api_hash) as client:
            сhannel_parts = link.split("t.me//")
            
            global channel_name
            channel_name = сhannel_parts[1]
            
            chat = await get_chat_info(channel_name, client)
            if not chat:
                await main()
                return
            
            print("Chat/channel found, start parsing...")
            await dump_users(chat, client)
            print(f"Chat/channel data saved in {json_file_path}")
            
            with open(json_file_path) as f:
                json_content = f.read()
                data = json.loads(json_content)
            
            print("Create .xlsx file...")
            create_excel_file(data)
            print(f"Path: {script_dir}\\excels\\{channel_name}.xlsx\nCreate .xlsx file: Done!")
            
            print("Create .txt file...")
            create_txt_file(data)
            print(f"Path: {script_dir}\\excels\\{channel_name}.txt\nCreate .txt file: Done!")
            
            # Закрываем?
            print("Press [Ctrl + C] to exit.")
            while True:
                pass
            
            
# ==================================================================================================== #

async def get_chat_info(username, client: TelegramClient):
    
    try:
        chat = await client(ResolveUsernameRequest(username))
    except:
        print('Chat/channel not found...\nRestarting...')
        asyncio.sleep(3)
        return False
    result = {
        'chat_id': chat.peer.channel_id,
        'access_hash': chat.chats[0].access_hash
    }
    return result


async def dump_users(chat, client: TelegramClient):
    offset = 0

    # Получаем участников чата
    chat_object = InputChannel(chat['chat_id'], chat['access_hash'])
    all_participants = []
    while True:
        
        participants = await client(GetParticipantsRequest(chat_object, ChannelParticipantsSearch(''), offset, limit, hash=0))
        if not participants.users:
            break
        
        all_participants.extend(participants.users)
        offset += len(participants.users)

    # Парсим тех, кто отправлял сообщения отдельно
    with open(config_file_path) as f:
        config = json.loads(f.read())
    
    limit_messages = config['limit_messages']
    bar_times = 0
    bar = progressbar.ProgressBar(maxval=len(all_participants) + limit_messages).start()
    
    active_participants = []
    async for message in client.iter_messages(chat_object, limit=limit_messages):
        bar_times += 1
        bar.update(bar_times) 
        sender_id = message.sender_id
        if sender_id not in active_participants:
            active_participants.append(sender_id)
    
    # Парсим всех и вся и записываем данные в json файл 
    all_user_details = []
    for participant in all_participants:
        bar_times += 1
        
        status = participant.status
        message = True if participant.id in active_participants else False
        
        if isinstance(status, UserStatusOffline):
            user_status = datetime.datetime.strftime(participant.status.was_online, '%d.%m.%Y %H:%M')
        elif isinstance(status, UserStatusOnline):
            user_status = datetime.datetime.strftime(participant.status.expires, '%d.%m.%Y %H:%M')
        elif isinstance(status, UserStatusRecently):
            user_status = 'Недавно'
        elif isinstance(status, UserStatusLastWeek):
            user_status = 'На прошлой неделе'
        elif isinstance(status, UserStatusLastMonth):
            user_status = 'В прошлом месяце'
        else:
            user_status = 'Не указано'
        
            
        info = {
                    "id": participant.id, 
                    "first_name": participant.first_name, 
                    "user": participant.username, 
                    "phone": participant.phone, 
                    "status": user_status,
                    "premium": participant.premium,
                    "message": message
                }
        
        all_user_details.append(info)    
        bar.update(bar_times)   
    
    with open(json_file_path, 'w') as outfile:
        json.dump(all_user_details, outfile)

    bar.finish()
    
# ==================================================================================================== #

def create_excel_file(data):
    book = openpyxl.Workbook()
    book.remove(book['Sheet'])
    
    create_all_sheets(book, data)
    
    book.save(f"excels\\{channel_name}.xlsx")
    book.close()


def create_txt_file(data):
    usernames = ''
    amount = 0
    for user in data:
        usernames += f"@{user['user']}\n"
        amount += 1
    
    with open(f"excels\\{channel_name}.txt", "w") as f:
        f.write(f"Удалось спарсить: {amount} участников.\n\n{usernames}")
    
# ==================================================================================================== #

def create_sheet(book: openpyxl.Workbook, data, sheet_name, condition=lambda user: True):
    """
    Создает лист в excel файле на основе данных, при этом применяя условие для фильтрации данных.

    Args:
        book: Объект Workbook из библиотеки openpyxl.
        data: Список словарей, содержащих данные для таблицы.
        sheet_name: Имя создаваемого листа.
        condition: Функция условия, которая принимает словарь пользователя и возвращает True, 
                    если пользователь должен быть добавлен в лист, иначе False. 
                    По умолчанию все пользователи добавляются.
    """
    sheet = book.create_sheet(sheet_name)
    set_headers(sheet)

    row = 2
    for user in data:
        if condition(user):
            set_values(sheet, row, user)
            row += 1

    formatting_cells(sheet)

# ==================================================================================================== #

def create_all_sheets(book, data):
    
    create_sheet(book, data, "Общие")
    create_sheet(book, data, "Премиум", lambda user: user["premium"])
    create_sheet(book, data, "Недавно (с временем)", condition=lambda user: user["status"] == "Недавно" or user["status"].count(".") == 2)
    create_sheet(book, data, "Недавно", lambda user: user["status"] == "Недавно")
    create_sheet(book, data, "С временем захода", lambda user: user["status"].count('.') == 2)
    create_sheet(book, data, "С номером", lambda user: user["phone"])
    create_sheet(book, data, "Отправляли сообщения", lambda user: user["message"])

# ==================================================================================================== #

def set_values(sheet, row, user):
    sheet[row][0].value = user['id'] 
    sheet[row][1].value = user['user'] 
    sheet[row][2].value = user['first_name']
    sheet[row][3].value = user['premium']  
    sheet[row][4].value = user['phone'] 
    sheet[row][5].value = user['status']
    sheet[row][6].value = user['message'] 

def set_headers(sheet):
    sheet['A1'] = 'User ID'
    sheet['B1'] = 'Username'
    sheet['C1'] = 'First name'
    sheet['D1'] = 'Premium'
    sheet['E1'] = 'Phone'
    sheet['F1'] = 'Status'
    sheet['G1'] = 'Message'

def formatting_cells(sheet):
    
    for col in sheet.columns:
        for cell in col:
            column_letter = cell.column_letter
            max_length = 0
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            sheet.column_dimensions[column_letter].width = max_length * 1.2

# ==================================================================================================== #

if __name__ == "__main__":
    asyncio.run(main())
