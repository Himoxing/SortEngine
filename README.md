POLISH and ENGILSH DOCUMENT Documentation and :

Polish/Polska:
Dokumentacja dla aplikacji SortEngine
Przegląd: Aplikacja SortEngine to proste narzędzie stworzone przy użyciu technologii HTML, VBScript oraz HTA (HTML Application). Zapewnia ona interfejs graficzny, dzięki któremu użytkownicy mogą wykonywać różne operacje na folderach, takie jak otwieranie folderów, sortowanie ich alfabetycznie oraz według daty ostatniej modyfikacji.
Funkcje:
    1. Otwórz Folder: Ta funkcja pozwala użytkownikom otworzyć okno dialogowe wyboru folderu i wybrać interesujący ich katalog. Aplikacja wyświetli listę podfolderów wybranego katalogu wraz z klikalnymi linkami, które umożliwiają otwarcie ich w eksploratorze plików systemu Windows.
    2. Sortuj Alfabetycznie: Ta funkcja sortuje listę podfolderów alfabetycznie i aktualizuje wyświetlanie odpowiednio. Użytkownicy mogą klikać na posortowane nazwy folderów, aby otworzyć je w eksploratorze plików systemu Windows.
    3. Sortuj Według Daty: Ta funkcja sortuje listę podfolderów na podstawie daty ostatniej modyfikacji, w kolejności rosnącej. Posortowane foldery są wyświetlane wraz z odpowiadającymi im datami ostatniej modyfikacji. Użytkownicy mogą klikać na posortowane nazwy folderów, aby otworzyć je w eksploratorze plików systemu Windows.
Interfejs Użytkownika: Interfejs użytkownika aplikacji składa się z głównego kontenera z następującymi komponentami:
    • Tytuł aplikacji: "SortEngine"
    • Trzy przyciski:
        1. "Otwórz folder/Open Folder": Rozpoczyna proces wyboru folderu i wyświetlania jego podfolderów.
        2. "Sortuj alfabetycznie/Sort Alphabet": Sortuje nazwy podfolderów alfabetycznie.
        3. "Sortuj według daty/Sort by Date": Sortuje podfoldery według daty ostatniej modyfikacji.
    • Dynamiczna lista folderów z klikalnymi linkami: Wyświetla nazwy podfolderów wraz z odpowiadającymi im linkami, które umożliwiają otwarcie ich.
Użycie:
    1. Kliknij przycisk "Otwórz folder/Open Folder", aby otworzyć okno dialogowe wyboru folderu. Wybierz folder z systemu.
    2. Aplikacja wyświetli listę podfolderów wybranego folderu jako klikalne linki.
    3. Kliknij przycisk "Sortuj alfabetycznie/Sort Alphabet", aby posortować listę podfolderów alfabetycznie.
    4. Kliknij przycisk "Sortuj według daty/Sort by Date", aby posortować listę podfolderów według daty ostatniej modyfikacji.
    5. Kliknij na dowolny link z nazwą podfolderu, aby otworzyć go w eksploratorze plików systemu Windows.
Szczegóły Techniczne:
    • Aplikacja jest tworzona przy użyciu HTML, VBScript oraz technologii HTA.
    • Funkcje VBScript są używane do interakcji z systemem, zarządzania operacjami na folderach oraz do wykonywania sortowania.
    • Ustawienia aplikacji HTA kontrolują wygląd oraz zachowanie okna aplikacji.
Uwaga: VBScript to starszy język skryptowy, który może posiadać ograniczenia oraz uwzględniać kwestie bezpieczeństwa. Funkcjonalność aplikacji może być wpływana zmianami w systemie operacyjnym lub w ustawieniach zabezpieczeń przeglądarek.
Ostrzeżenie: Niniejsza dokumentacja dostarcza ogólnego przeglądu funkcji oraz sposobu użytkowania aplikacji SortEngine. Użytkownicy powinni mieć świadomość potencjalnych zagrożeń związanych z używaniem aplikacji HTA oraz języka VBScript w nowoczesnych środowiskach. Zawsze zachowuj ostrożność podczas uruchamiania skryptów lub aplikacji z niezaufanych źródeł.


English/Angielska:

Documentation for SortEngine Application
Overview: The SortEngine application is a simple utility developed using HTML, VBScript, and the HTA (HTML Application) technology. It provides a graphical user interface for users to perform various operations on folders, such as opening folders, sorting folders alphabetically, and sorting folders by their last modified date.
Features:
    1. Open Folder: This feature allows users to open a folder dialog and select a folder. The application then lists the subfolders of the selected folder along with clickable links to open them in Windows File Explorer.
    2. Sort Alphabetically: This feature sorts the list of subfolders alphabetically and updates the display accordingly. Users can click on the sorted folder names to open them in Windows File Explorer.
    3. Sort by Date: This feature sorts the list of subfolders based on their last modified date in ascending order. The sorted folders are displayed along with their corresponding last modified dates. Users can click on the sorted folder names to open them in Windows File Explorer.
User Interface: The application's user interface consists of a main container with the following components:
    • Application title: "SortEngine"
    • Three buttons:
        1. "Otwórz folder/Open Folder": Initiates the process of selecting a folder and listing its subfolders.
        2. "Sortuj alfabetycznie/Sort Alphabet": Sorts the subfolders' names in alphabetical order.
        3. "Sortuj według daty/Sort by Date": Sorts the subfolders based on their last modified date.
    • A dynamic list of folders with clickable links: Displays the names of subfolders along with corresponding links that allow users to open them.
Usage:
    1. Click the "Otwórz folder/Open Folder" button to open a folder dialog. Select a folder from your system.
    2. The application will display the list of subfolders within the selected folder as clickable links.
    3. Click the "Sortuj alfabetycznie/Sort Alphabet" button to sort the list of subfolders alphabetically.
    4. Click the "Sortuj według daty/Sort by Date" button to sort the list of subfolders based on their last modified date.
    5. Click on any subfolder name link to open that folder in Windows File Explorer.
Technical Details:
    • The application is developed using HTML, VBScript, and HTA technology.
    • VBScript functions are used to interact with the system, manage folder operations, and perform sorting.
    • The HTA application settings control the appearance and behavior of the application window.
Note: VBScript is an older scripting language that may have limitations and security considerations. The application's functionality may be affected by changes in the operating system or browser security settings.
Disclaimer: This documentation provides an overview of the SortEngine application's features and usage. Users should be aware of potential security risks associated with using HTA applications and VBScript in modern environments. Always exercise caution when running scripts or applications from untrusted sources.

Polish:

przejdźmy teraz przez każdą z funkcji znajdujących się w kodzie aplikacji SortEngine i wyjaśnijmy, co dokładnie robią:
    1. OpenFolderDialog:
        ◦ Ta funkcja jest wywoływana po kliknięciu przycisku "Otwórz folder/Open Folder" i ma na celu otwarcie okna dialogowego do wyboru folderu.
        ◦ Tworzy się instancję obiektu Shell.Application, który reprezentuje interfejs dostępu do funkcji eksploratora systemu.
        ◦ Użytkownik wybiera folder, a wybrany folder jest reprezentowany jako obiekt selectedFolder.
        ◦ Następnie tworzona jest instancja obiektu Scripting.FileSystemObject (fs), który umożliwia operacje na plikach i folderach.
        ◦ Wszystkie podfoldery wybranego folderu są zapisywane w zmiennej subfolders.
        ◦ Iterujemy przez każdy z podfolderów (subfolder) i tworzymy rekordy zawierające nazwę, ścieżkę i datę ostatniej modyfikacji tych podfolderów.
        ◦ Zmienne te (folderNamesArray, folderPaths, folderDates) są przechowywane jako tablice indeksowane.
        ◦ Wszystkie nazwy podfolderów są agregowane jako łącze HTML do zmiennej folderNames.
        ◦ Zawartość zmiennej folderNames jest przypisywana do elementu z identyfikatorem "folderList" w dokumencie HTML.
    2. OpenFolder(index):
        ◦ Ta funkcja jest wywoływana po kliknięciu na nazwę podfolderu w liście. Otwiera ona wybrany folder w eksploratorze plików systemu Windows.
        ◦ Tworzy się instancję obiektu WScript.Shell, który umożliwia uruchamianie poleceń systemowych.
        ◦ Używając metody Run obiektu WScript.Shell, uruchamiamy eksplorator plików i przekazujemy mu ścieżkę do wybranego folderu.
    3. SortFoldersAlphabetically:
        ◦ Ta funkcja jest wywoływana po kliknięciu przycisku "Sortuj alfabetycznie/Sort Alphabet" i służy do sortowania nazw podfolderów alfabetycznie.
        ◦ Tworzymy instancję obiektu System.Collections.ArrayList, który pozwala na dynamiczne dodawanie i sortowanie elementów.
        ◦ Iterujemy przez nazwy podfolderów (folderName) i dodajemy je do obiektu ArrayList.
        ◦ Sortujemy zawartość obiektu ArrayList za pomocą metody Sort.
        ◦ Iterujemy ponownie przez posortowane nazwy podfolderów i odnajdujemy ich indeksy w oryginalnej tablicy. Tworzymy łącza HTML z posortowanymi nazwami i przypisujemy je do elementu "folderList".
    4. SortFoldersByDate:
        ◦ Ta funkcja jest wywoływana po kliknięciu przycisku "Sortuj według daty/Sort by Date" i służy do sortowania podfolderów według daty ostatniej modyfikacji.
        ◦ Za pomocą dwóch pętli for iterujemy przez tablicę dat modyfikacji podfolderów i porównujemy je ze sobą.
        ◦ Jeśli data w podfolderze j jest wcześniejsza niż w podfolderze i, zamieniamy dane i nazwy miejscami.
        ◦ Tworzymy łącza HTML z posortowanymi nazwami podfolderów oraz ich datami i przypisujemy je do elementu "folderList".
Uwagi końcowe: Każda z funkcji wykonuje konkretne operacje związane z wyświetlaniem, sortowaniem i otwieraniem folderów. Wspólnie tworzą one interaktywny interfejs użytkownika, umożliwiający wybieranie folderów oraz sortowanie ich zawartości w różny sposób.


English:

let's explain each of the functions present in the SortEngine code, detailing what they do:
    1. OpenFolderDialog:
        ◦ This function is called when the "Otwórz folder/Open Folder" button is clicked. Its purpose is to open a folder selection dialog.
        ◦ An instance of the Shell.Application object is created, representing the interface for accessing system explorer functionalities.
        ◦ The user selects a folder, and the selected folder is represented as the selectedFolder object.
        ◦ Then, an instance of the Scripting.FileSystemObject (fs) is created, enabling operations on files and folders.
        ◦ All subfolders of the selected folder are stored in the subfolders variable.
        ◦ Each subfolder (subfolder) is iterated over, and records containing the name, path, and last modification date of these subfolders are created.
        ◦ These variables (folderNamesArray, folderPaths, folderDates) are stored as indexed arrays.
        ◦ All subfolder names are aggregated as HTML links into the folderNames variable.
        ◦ The content of the folderNames variable is assigned to the element with the ID "folderList" in the HTML document.
    2. OpenFolder(index):
        ◦ This function is called when a subfolder name in the list is clicked. It opens the selected folder in the Windows file explorer.
        ◦ An instance of the WScript.Shell object is created, allowing the execution of system commands.
        ◦ Using the Run method of the WScript.Shell object, we run the file explorer and pass the path to the selected folder.
    3. SortFoldersAlphabetically:
        ◦ This function is called when the "Sortuj alfabetycznie/Sort Alphabet" button is clicked and is used to sort subfolder names alphabetically.
        ◦ We create an instance of the System.Collections.ArrayList object, which allows dynamic addition and sorting of elements.
        ◦ We iterate through subfolder names (folderName) and add them to the ArrayList object.
        ◦ We sort the content of the ArrayList using the Sort method.
        ◦ We iterate again through the sorted subfolder names, find their indices in the original array, create HTML links with sorted names, and assign them to the "folderList" element.
    4. SortFoldersByDate:
        ◦ This function is called when the "Sortuj według daty/Sort by Date" button is clicked and is used to sort subfolders by their last modification date.
        ◦ Using two for loops, we iterate through the array of modification dates of subfolders and compare them.
        ◦ If the date in subfolder j is earlier than in subfolder i, we swap the data and names.
        ◦ We create HTML links with sorted subfolder names and their dates and assign them to the "folderList" element.
Closing Notes: Each function performs specific operations related to displaying, sorting, and opening folders. Collectively, they create an interactive user interface that allows folder selection and sorting of their contents in different ways.

