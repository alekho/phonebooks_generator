# Генератор адресной книги

Самописный скрипт на **python** для генерации адресной книги из файла **xlsx**. Для телефонов Yealink и Eltex.
Так же сделал **exe** для удобства использования.
В репозитории лежит файл **phonebook.xlsx**, с примером заполнения. Важное уточнение, между строчками не должно быть пробелов(пустых строк). Иначе не сработает.

В результате работы программы Вы получите два файла формата xml.

**Для Eltex, с содержимым:**

```xml
<?xml version='1.0' encoding='UTF-8'?>
<EltexIPPhoneDirectory>
    <Title>Phones</Title>
    <Prompt>Prompt</Prompt>
    <Grouplist>
        <Group name="отдел1"/>
        <Group name="отдел2"/>
        <Group name="отдел3"/>
    </Grouplist>
    <DirectoryEntry>
        <Name>Иванов</Name>
        <Telephone>1111</Telephone>
        <Telephone>8987654321</Telephone>
        <Telephone>3333</Telephone>
        <Group>отдел1</Group>
    </DirectoryEntry>
    <DirectoryEntry>
        <Name>Петров</Name>
        <Telephone>2222</Telephone>
        <Telephone>4444</Telephone>
        <Group>отдел2</Group>
    </DirectoryEntry>
    <DirectoryEntry>
        <Name>Сидоров</Name>
        <Telephone>3333</Telephone>
        <Group>отдел3</Group>
    </DirectoryEntry>
    <DirectoryEntry>
        <Name>Обломов</Name>
        <Telephone>4444</Telephone>
        <Group>отдел1</Group>
        </DirectoryEntry>
    </EltexIPPhoneDirectory>
```

**Для Yealink, с содержимым:**
```xml
<?xml version='1.0' encoding='UTF-8'?>
<YealinkIPPhoneBook>
    <Title>Yealink</Title>
    <Menu Name="отдел1">
        <Unit Name="Иванов" Phone1="1111" Phone2="8987654321" Phone3="3333" default_photo="Resource:"/>
        <Unit Name="Обломов" Phone1="4444" Phone2="" Phone3="" default_photo="Resource:"/>
    </Menu>
    <Menu Name="отдел2">
        <Unit Name="Петров" Phone1="2222" Phone2="" Phone3="4444" default_photo="Resource:"/>
    </Menu>
    <Menu Name="отдел3">
        <Unit Name="Сидоров" Phone1="3333" Phone2="" Phone3="" default_photo="Resource:"/>
    </Menu>
</YealinkIPPhoneBook>
```

Данный скрипт выложил "как есть", программистом не являюсь. Скорее всего можно сделать аккуратнее и правильнее =)

 