


function sendMap() {
    const sh = SpreadsheetApp.getActiveSheet();
    const address = sh.getRange('A1').getValue()
    const map = Maps.newStaticMap().addMarker(address);
    GmailApp.sendEmail('dmaberlin77@gmail.com', 'Map', 'see bottom', {
        attachments: [map],
    })
}


//@customfunction

function USDTOUAH(USD) {
    var cache = CacheService.getScriptCache();

    var rate = cache.get('rates.UAH')

    if (!rate) {
        var response = UrlFetchApp.fetch(
            "https://api.exchangeratesapi.io/latest?base=USD"
        );

        var result = JSON.parse(response.getContentText());
        rate = result.rates.UAH
        cache.put('rates.UAH', rate)
    }
    var UAH = USD * Number(rate);
    return UAH + ' ₴'
}






//plug
function func2() {

}
//plug
function func3() {

}




function Myfunc() {
    var spreadsheet = SpreadsheetApp.getActive()
    var sheet = spreadsheet.getActiveSheet();
    sheet.getRange(spreadsheet.getCurrentCell().getRow(), 1, 1, sheet.getMaxColumns()).activate()
    spreadsheet.getActiveRangeList().setBackground('#4c1130')
        .setBorder('#ffffff')
        .setFontWeight('bold');
    spreadsheet.getActiveSheet().setFrozenRows(1);

}



function addToDatabase0() {
    var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1gsPEw8zoP5nuWnq4EAkaY62Ympy1BGb069BP2VhGBXg/edit?usp=sharing');
    var sh_db_movies = ss.getSheetByName('dbmovie');
    var values = sh_db_movies.getRange(2, 1, 11, 11).getValues()
    Logger.log(values)
}

function addToDatabase() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dbmovie')
    var ss_active = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('db_active')
    var values = ss_active.getRange(2, 1, 11, 11).getValues()

    Logger.log(values)

}

function addHeaders() {
    var ss_db_movies = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dbmovie')
    var ss_active = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('db_active')
    var headers = ss_db_movies.getSheetValues(1, 1, 1, 11)
    Logger.log(headers)
}






// Получает лист, диапазон данных и значения
// Таблицы, хранящейся в ss_movie

function loadMovieList() {
    var sheet = SpreadsheetApp.getActiveSheet()
    var ss_movie = SpreadsheetApp.openById('1oLoTej9MpBXR6-JE0bP_WouDXpHo14AEFYGXdaMDuq4')
    var sheet_movie = ss_movie.getSheetByName('db_relevant')
    var movie_range = sheet_movie.getDataRange();
    var movie_list_values = movie_range.getValues();

    const date = new Date()
    const options1 = {
        hour: 'numeric',
        minute: 'numeric',
        month: 'short',
        weekday: 'short',
        // year:'2-digit'
    }



    // var db_name_updated=new Intl.DateTimeFormat('ru-RU',options1).format(date)
    var db_name_updated = new Intl.DateTimeFormat('en-US', options1).format(date)



    // Добавляет эти значения в активный лист
    // Текущей Таблицы. Это действие перезаписывает все данные, какие уже
    // были до момента вызова функции. 
    sheet.getRange(1, 1, movie_range.getHeight(), movie_range.getWidth()).setValues(movie_list_values)

    // rename list  и изменяем размеры колонок

    sheet.setName(db_name_updated);
    sheet.autoResizeColumns(1, 3);

}


////splits the author and title into their respective cells when commas are detected

/**
 * Changes the header and author columns,
 * splits the value of the header column by the first comma, if any.
 */
function splitAtFirstComma() {
    // Получает активный (в данный момент выделенный) диапазон.
    var active_range = SpreadsheetApp.getActiveRange();
    var title_author_range = active_range.offset(0, 0, active_range.getHeight(), active_range.getWidth() + 1);

    // Получает значения выбранных ячеек.
    // Это двумерный массив.
    var title_author_values = title_author_range.getValues();


    // Обновляем значения там, где есть запятые.
    // Предполагается, что наличие запятой указывает на шаблон "авторы, название".
    for (var row = 0; row < title_author_values.length; row++) {
        var index_of_first_comma = title_author_values[row][0].indexOf(', ');

        if (index_of_first_comma >= 0) {
            // Найдена запятая, разделяет и обновляет значения в массиве значений.
            var titles_and_authors = title_author_values[row][0];

            // Обновляет значение заголовка в массиве.
            title_author_values[row][0] = titles_and_authors.slice(index_of_first_comma + 2)

            //Обновляет значение автора в массиве.
            title_author_values[row][1] = titles_and_authors.slice(0, index_of_first_comma);
        }
    }

    //Помещает обновленные значения обратно в таблицу.
    title_author_range.setValues(title_author_values);
}



/** 
 * Changes the header and author columns,
 * separating the value of the header column by the occurrence of the word " by ", if present.
 */

function splitAtLastBy() {
    // Gets the active (currently selected) range.
    var activeRange = SpreadsheetApp.getActiveRange();
    var title_author_range = activeRange.offset(0, 0, activeRange.getHeight(), activeRange.getWidth() + 1);

    // Получает значения выбранных ячеек.
    // Это двумерный массив.
    var title_author_values = title_author_range.getValues();

    // Обновляем значения там, где есть текст " by ".
    // Предполагается, что наличие 'by' указывает на шаблон "название by авторы"
    for (var row = 0; row < title_author_values.length; row++) {
        var index_of_last_by = title_author_values[row][0].lastIndexOf(' by ');

        if (index_of_last_by >= 0) {
            //Найдена фраза 'by',разделяет и обновляет значения в массиве значений.
            var titles_and_authors = title_author_values[row][0];

            // Обновляет значение заголовка в массиве.
            title_author_values[row][0] = titles_and_authors.slice(0, index_of_last_by);

            //обновляет значение автора в массиве.
            title_author_values[row][1] = titles_and_authors.slice(index_of_last_by + 4);
        }
    }

    // Помещает обновленные значения обратно в Таблицу.
    title_author_range.setValues(title_author_values)

}





/**
 *
* Helper function,
 * which retrieves book data from the publicly available Open Library API.
 *
 * @param {number} ISBN - ISBN number of the book you want to find.
 * @return {object} Book data in JSON format.
 */

var BASE_URL_BOOK = "https://openlibrary.org/api/books?bibkeys=ISBN:"
var PARAMS_BOOK = '&jscmd=details&format=json'

// &jscmd=details&format=json - это строка параметров, которая может добавляться в URL адрес веб-страницы. Эта строка используется для передачи параметров запроса на сервер.

function fetchBookData_(ISBN) {
    // Connection to the public API.

    var URL = BASE_URL_BOOK + ISBN + PARAMS_BOOK;
    var response = UrlFetchApp.fetch(URL, { 'muteHttpExceptions': true });

    // Делает запрос к API и получает ответ.
    var json = response.getContentText();
    var book_data = JSON.parse(json)


    // Возвращает только интересующую нас информацию. 
    return book_data['ISBN:' + ISBN]
}




/**
 * Fills in missing header and author data
 * With Open Library API calls.
 */

function fillInTheBlanks() {
    // Константы, определяющие индекс столбцов заголовка, автора и ISBN
    // (в двумерном массиве bookValues ниже).
    var TITLE_COLUMN = 0;
    var AUTHOR_COLUMN = 1;
    var ISBN_COLUMN = 2;

    // Получает информацию о текущей книге на активном листе.
    // Данные помещаются в двумерный массив.
    var data_range = SpreadsheetApp.getActiveSpreadsheet().getDataRange();
    var book_values = data_range.getValues();


    // Проверяет каждую строку данных (исключая строку заголовка).
    // Если ISBN присутствует, а заголовок или автор отсутствуют,
    // используется метод fetchBookData_(ISBN)
    // для получения недостающих данных из Open Library API.
    // Заполняет недостающие названия или авторов, когда они будут найдены.

    for (var row = 1; row < book_values.length; row++) {
        var isbn = book_values[row][ISBN_COLUMN];
        var title = book_values[row][TITLE_COLUMN];
        var author = book_values[row][AUTHOR_COLUMN];


        if (isbn != '' && (title === '' || author === '')) {
            var book_data = fetchBookData_(isbn)


            // Иногда API не возвращает необходимую информацию.
            // В таких случаях не пытается обновить строку дальше.
            if (!book_data || !book_data.details) continue;

            // API может не иметь заголовка, поэтому заполняет его, только если API возвращает заголовок, а заголовок на листе пуст.
            if (title === '' && book_data.details.title) {
                book_values[row][TITLE_COLUMN] = book_data.details.title;
            }

            // API может не иметь имени автора, поэтому заполняет его, только в том случае, если API возвращает автора, а автор пуст в таблице.

            if (author === '' && book_data.details.authors && book_data.details.authors[0].name) {
                book_values[row][AUTHOR_COLUMN] = book_data.details.author[0].name
            }
        }
    }
    // puts the updated book data values back into the Table.
    data_range.setValues(book_values);

}



function formatRowHeader(){
    var sheet=SpreadsheetApp.getActiveSheet();
    var header_range=sheet.getRange(1,1,1,sheet.getLastColumn());

// Applies a specific format to the top line:
  // bold white text,
  // blue-green background
  // and a solid black border around the cells.

  header_range
  .setFontWeight('bold')
  .setFontColor('#ffffff')
  .setBackground('#007272')
  .setBorder(true,true,true,true,null,null,null,SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

}

/**
 * Formats the name column of the active sheet.
 */

function formatColumnHeader(){
    var sheet=SpreadsheetApp.getActiveSheet();

    //get total numbers lines on the data range, not including header line.
    var num_rows=sheet.getDataRange().getLastRow()-1;

    /** 
     *  метод getLastRow() используется для получения номера последней строки в этом диапазоне. Однако, поскольку строка заголовка обычно не считается частью данных, то вычитается 1, чтобы получить общее количество строк в диапазоне без учета строки заголовка.

Таким образом, переменная numRows будет содержать общее количество строк в диапазоне данных, не включая строку заголовка. Это значение может быть использовано для различных целей, например, для создания цикла для обработки всех строк в диапазоне данных.
    */

    //get range with names
    var column_header_range=sheet.getRange(2,1,num_rows,1);

    //Applies text formatting and adds borders;
    column_header_range
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setBorder(true,true,true,true,null,null,null,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


    //// Calls an auxiliary method to create a hyperlink
  // in first column, based on data of 'Link' column

  hyperlinkColumnHeaders_(column_header_range,num_rows);

}



/**
* Helper function that links the title column
 * with the contents of the "Link" column.
 * The function then removes the "Link" column.
 *
 * @param {object} header_range Range with names to update.
 * @param {number} num_rows Size of range with names.
 */

  function hyperlinkColumnHeaders_(header_range,num_rows) {
    var url_title='Link'
    //Gets column indexes witn names and references
    var header_col_index=1;
    var url_col_index=columnIndexOf_(url_title)
    
    //return if no url column
    if(url_col_index ==-1) return;

    //takes values from name column and url column
        var url_range=header_range.offset(0,url_col_index-header_col_index);
        var header_values=header_range.getValues();
        var url_values=url_range.getValues();


    //Update title values to hyperlinked values
    for(var row=0;row<num_rows;row++){
        header_values[row][0]='=HYPERLINK("'+url_values[row]+'","'+header_values[row]+'")';
    }
    header_range.setValues(header_values);


    //Remove the column with links, not clutter the sheet with unnecessary data
    SpreadsheetApp.getActiveSheet().deleteColumn(url_col_index);
  }

/**
* Helper function that looks through all column headers
* and returns the index of the column by name on line 1.
* If a column with that name does not exist, this function returns -1.
* If more than one column has the same name on line 1,
* the index of the first column detected is returned.
*
* @param {string} colName The name to look for in the column headers.
* @return The index of the column on the active sheet or -1 if no name is found.







//**************************************UI ***********************************************/
//**************************************UI ***********************************************/
//**************************************UI ***********************************************/




// var ss=SpreadsheetApp.getActiveSpreadsheet()
// var sheet_db_active=ss.getSheetByName('db_active')

/**
 * A special function performed when the Table 
 * has opened or reloaded. onOpen() is used to add
 * Custom menus.
 */



function onOpen() {
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('Fast format')
    .addItem('Format header row','formatRowHeader')
    .addItem('Format title column','formatColumnHeader')
    .addItem('Format data set','formatDataset')
    .addToUi();

    ui.createMenu('Custom Menu')
        .addItem('currency exchange', 'usdToUAH')
        .addSeparator()
        .addItem('load Movie list', 'loadMovieList')
        .addToUi()

    ui.createMenu('Border Menu')
        .addItem('get border', 'getBorder')
        .addSeparator()
        .addItem('remove border ', 'removeBorder')
        .addToUi();

    ui.createMenu('Book List')
        .addItem('upload book list', 'loadBookList')
        .addSeparator()
        .addItem('Separate title/author by first comma', 'splitAtFirstComma')
        .addItem('Separate title/author by last \'by\'', 'splitAtLastBy')
        .addSeparator()
        .addItem('Fill in empty title and author cells', 'fillInTheBlanks')
        .addToUi()


    ui.createMenu('onCheck')
        .addItem('add chechbox', 'onCheckButtons')
        .addToUi()
}

function onEdit() {
    getBorder()

}



// ****************** button check****************

function onCheckButtons() {
    // var lr=sheet.getLastRow();
    // var range=sheet.getRange(1,lc,1,lc+2)


    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var lc = sheet.getLastColumn();

    var toggleBorderonEdit = sheet.getRange(1, lc + 2);


    // cell.setDataValidations(undefined)

    toggleBorderonEdit.insertCheckboxes()
    var endColumn = toggleBorderonEdit.getColumn()
    // var toggleBorderonEditInfo=sheet.getRange(2,lc+2,3,endColumn)
    var toggleBorderonEditInfo = sheet.getRange(2, endColumn)

    toggleBorderonEditInfo.merge().setValue(CHECKBOX_1)
    // toggleBorderonEditInfo.setValue(CHECKBOX_1)




}



//**************************Border***********************

function getBorder() {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    // var sheet=ss.getSheetByName('db_active')
    var sheet = ss.getActiveSheet()
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    sheet.getRange('A:P').setBorder(false, false, false, false, false, false, '#000', SpreadsheetApp.BorderStyle.DOUBLE)
    sheet.getRange(1, 1, lr, lc).setBorder(true, true, true, true, true, true, '#000', SpreadsheetApp.BorderStyle.SOLID)

}

function removeBorder() {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = ss.getActiveSheet()
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    sheet.getRange(1, 1, lr + 1, lc + 1).setBorder(false, false, false, false, false, false, '#000', SpreadsheetApp.BorderStyle.DOUBLE)
    // sheet.getRange(1,1,lr,lc).setBorder(true,true,true,true,true,true,'#000',SpreadsheetApp.BorderStyle.SOLID)
}