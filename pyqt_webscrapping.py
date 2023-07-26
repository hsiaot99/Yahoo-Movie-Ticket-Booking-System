import io
import json
import re
import sys
import datetime
from pathlib import Path

import folium
import geocoder
import pandas as pd
import requests
from PyQt6 import uic
from PyQt6.QtCore import Qt, QAbstractTableModel
from PyQt6.QtGui import QColor, QIcon, QPixmap
from PyQt6.QtWidgets import QMainWindow, QApplication, QMessageBox, QAbstractItemView, QHeaderView
from bs4 import BeautifulSoup


STYLE_SHEET = """
    QMainWindow {
        border-image: url(background.png); 
    }
    
    #title {
        background-color: rbg(100, 100, 100);
        color: white;
    }
    
    #stackedWidget {
        background: white;
    }
    
    #btn_all_movie, #btn_weekly_movie {
        background-color: rbg(100, 100, 100);
        border: 1px solid blue;
        color: blue;
    }
    
    #btn_all_movie:hover, #btn_weekly_movie:hover {
        background-color: blue;
        color: white;
    }
    
    #btn_all_movie:pressed, #btn_weekly_movie:pressed {
        background-color: white;
        color: blue;
    }
    
    #pTextEdit_intro {
        background-color: transparent;
    }
"""


class MainWindow(QMainWindow):
    UI = r'./pyqt_webscrapping.ui'
    ICON = r'./popcorn.png'

    def __init__(self, *args, **kwargs):
        # Initialize UI
        super(MainWindow, self).__init__(*args, **kwargs)
        uic.loadUi(self.UI, self)
        self.setWindowTitle('PyQt - WebScrapping')
        self.setWindowIcon(QIcon(self.ICON))
        self.icon.setPixmap(QPixmap(self.ICON))

        # Set up tableview
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)  # row selection
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)  # single row select
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

        # Initialize parameters
        self.yahoo_movie = None
        self.df = None

        # Set up ui signals.
        self.frame_btns.hide()
        self.btn_all_movie.clicked.connect(self.display_all_movie)
        self.btn_weekly_movie.clicked.connect(self.display_weekly_movie)
        self.btn_movie_detail.clicked.connect(self.display_movie_detail)
        self.btn_back.clicked.connect(self.display_back)
        # self.btn_show_time.clicked.connect(self.display_map)

        # Set the first page.
        self.stackedWidget.setCurrentIndex(0)

        self.setStyleSheet(STYLE_SHEET)

    def display_loading_page(self):
        QApplication.processEvents()

        # Load data.
        self.yahoo_movie = YahooMovie()

        # Switch to next page.
        self.frame_btns.show()
        self.display_all_movie()

    def display_all_movie(self):
        self.stackedWidget.setCurrentIndex(1)

        # Set up model
        self.df = self.yahoo_movie.query_movie()
        model = TableModel(self.df)
        self.table.setModel(model)
        for column_hidden in [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]:  # Hide some columns (Abstract, PaperText, imgfile)
            self.table.hideColumn(column_hidden)

        # Connect current model with event (To display detail)
        selection_model = self.table.selectionModel()
        selection_model.selectionChanged.connect(self.selection_changed)

        index = model.index(0, 1)
        self.table.setCurrentIndex(index)

    def display_weekly_movie(self):
        self.stackedWidget.setCurrentIndex(1)

        # Set up model
        self.df = self.yahoo_movie.query_movie(True)
        model = TableModel(self.df)
        self.table.setModel(model)
        for column_hidden in [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]:  # Hide some columns (Abstract, PaperText, imgfile)
            self.table.hideColumn(column_hidden)

        # Connect current model with event (To display detail)
        selection_model = self.table.selectionModel()
        selection_model.selectionChanged.connect(self.selection_changed)

        index = model.index(0, 1)
        self.table.setCurrentIndex(index)

    def display_movie_detail(self):
        self.stackedWidget.setCurrentIndex(2)

        row = self.table.currentIndex().row() + 1

        self.detail_movie_name.setText(str(self.df.loc[row]['電影名稱'])+'  '+str(self.df.loc[row]['英文名稱']))
        self.detail_release_time.setText(str(self.df.loc[row]['上映日期']))
        self.detail_run_time.setText(str(self.df.loc[row]['片長']))
        self.detail_expectation.setText(str(self.df.loc[row]['期待度']))
        self.detail_satisfaction.setText(str(self.df.loc[row]['滿意度']))
        self.detail_imdb.setText(str(self.df.loc[row]['IMDb分數']))
        self.detail_company.setText(str(self.df.loc[row]['發行公司']))
        self.detail_director.setText(str(self.df.loc[row]['導演']))
        self.detail_cast.setText(str(self.df.loc[row]['演員']))
        self.detail_intro.setPlainText(str(self.df.loc[row]['劇情介紹']))

        pixmap = QPixmap()
        pixmap.loadFromData(requests.get(str(self.df.loc[row]['電影海報'])).content)
        self.detail_img.setPixmap(pixmap)

    def display_back(self):
        self.stackedWidget.setCurrentIndex(1)

    def selection_changed(self, selected, deselected):
        current_index = selected.indexes()[0]  # single selection
        row = current_index.row() + 1

        self.label_movie_name.setText(str(self.df.loc[row]['電影名稱'])+'<br>'+str(self.df.loc[row]['英文名稱']))
        self.label_release_time.setText(str(self.df.loc[row]['上映日期']))
        self.label_imdb.setText(str(self.df.loc[row]['IMDb分數']))

        intro = str(self.df.loc[row]['劇情介紹'])
        if len(intro) > 200:
            intro = intro[:200] + '...'
        self.label_intro.setText(intro)

        pixmap = QPixmap()
        pixmap.loadFromData(requests.get(str(self.df.loc[row]['電影海報'])).content)
        self.label_img.setPixmap(pixmap)

    def display_map(self):  # TODO
        row = self.table.currentIndex().row() + 1
        movie_id = self.df.loc[row]['電影ID']
        print(movie_id)

        # Get movie time.
        today = datetime.date.today()
        show_times = self.yahoo_movie.get_movie_time(movie_id, str(today))
        # for idx in range(0, 7):
        #     date = today + datetime.timedelta(days=idx)
        #     self.yahoo_movie.get_movie_time(movie_id, str(date))
        #
        print(show_times)

        theater_id = show_times['戲院ID'].unique()[0]
        print(theater_id)

        theaters = self.yahoo_movie.theaters
        theaters = theaters.loc[theaters['戲院ID'] == theater_id]
        print(theaters)

        m = folium.Map(
            tiles='Stamen Terrain',
            zoom_start=13,
            location=[theaters[0]['緯度'], theaters[0]['經度']]
        )
        print('OK')

        # save map data to data object
        data = io.BytesIO()
        folium.Marker(location=[theaters[0]['緯度'], theaters[0]['經度']]).add_to(m)  # 插入圖標
        m.save(data, close_file=False)
        print('OK')

        self.webview.setHtml(data.getvalue().decode())

    def closeEvent(self, event):
        # Create a message box.
        messagebox = QMessageBox()
        messagebox.setWindowTitle('Message')
        messagebox.setText('Are you sure you want to exit the dialog?')
        messagebox.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        # If user click "Yes", close the app.
        reply = messagebox.exec()
        if reply == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]  # pandas' iloc method
            return str(value)

        if role == Qt.ItemDataRole.TextAlignmentRole:
            return Qt.AlignmentFlag.AlignVCenter + Qt.AlignmentFlag.AlignHCenter

        if role == Qt.ItemDataRole.BackgroundRole and (index.row() % 2 == 0):
            return QColor('#A6CDE7')

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]

    # Add Row and Column header
    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.ItemDataRole.DisplayRole:  # more roles
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])

            if orientation == Qt.Orientation.Vertical:
                return str(self._data.index[section])


class YahooMovie:
    DATA_DIR = Path('./data')
    MOVIE_FILE = DATA_DIR / 'movies.xlsx'
    THEATER_FILE = DATA_DIR / 'theaters.xlsx'

    def __init__(self):
        self.movies = pd.DataFrame(
            dtype=str,
            columns=['電影ID', '電影名稱', '英文名稱', '上映日期', '片長',
                     '期待度', '滿意度', 'IMDb分數', '發行公司', '導演',
                     '演員', '電影海報', '劇情介紹']
        )
        self.show_times = pd.DataFrame(dtype=str, columns=['電影ID', '戲院ID', '類型', '日期', '時間'])
        self.theaters = pd.DataFrame(dtype=str, columns=['戲院ID', '戲院名稱', '地區', '電話', '地址', '緯度', '經度'])
        self.theater_areas = [
            '臺北', '新北', '基隆', '桃園', '新竹', '苗栗', '臺中', '彰化', '南投', '雲林',
            '嘉義', '臺南', '高雄', '屏東', '宜蘭', '花蓮', '臺東', '澎湖', '金門', '連江',
        ]

        self.DATA_DIR.mkdir(exist_ok=True, parents=True)
        self.load_movies()
        self.load_theaters()
        self.download_all_movies()

    def load_movies(self):
        if self.MOVIE_FILE.exists():
            try:
                self.movies = pd.read_excel(self.MOVIE_FILE, dtype=str)
            except:
                print('無法載入之前的電影資料')

    def save_movies(self):
        self.movies = self.movies.astype(str)
        self.movies.to_excel(self.MOVIE_FILE, index=False)

    def load_theaters(self):
        if self.THEATER_FILE.exists():
            try:
                self.theaters = pd.read_excel(self.THEATER_FILE, dtype=str)
            except:
                print('無法載入之前的戲院資料')

    def save_theaters(self):
        self.theaters = self.theaters.astype(str)
        self.theaters.to_excel(self.THEATER_FILE, index=False)

    def download_all_movies(self):
        api = 'https://movies.yahoo.com.tw/movie_intheaters.html?page='

        for page in range(1, 1000):
            url = api + str(page)
            res = requests.get(url)
            print(url)  # TODO
            soup = BeautifulSoup(res.text, 'html.parser')

            # Check if there is no movie in this page.
            page_is_empty = soup.find('p', string='本週無電影/戲劇上映。')
            if page_is_empty:
                break

            # Get movie list.
            movie_elements = soup.find('ul', class_='release_list').find_all('li')
            for element in movie_elements:
                # There are two big <div> in each <li>.
                div_release_foto = element.find('div', class_='release_foto')
                div_release_info = element.find('div', class_='release_info')
                movie_time_url = div_release_info.find('a', string=re.compile('時刻表'))['href']
                movie_id = movie_time_url.split('id=')[1]

                # Get movie info if it does not be saved before.
                if movie_id not in self.movies['電影ID'].unique():
                    # Get movie info.
                    chinese_name = div_release_info.find('div', class_='release_movie_name').a.text.strip()
                    print(f'新增電影資訊: {chinese_name}')

                    english_name = div_release_info.find('div', class_='en').a.text.strip()
                    expectation = div_release_info.find('div', class_='level_name', string='期待度').find_next_sibling('div').span.text.strip()
                    satisfaction = div_release_info.find('div', class_='level_name', string='滿意度').find_next_sibling('div').span['data-num']
                    release_time = div_release_info.find('div', class_='release_movie_time').text.replace('上映日期：', '').strip()
                    intro = div_release_info.find('div', class_='release_text').span.text.strip()
                    img_url = div_release_foto.a.img['data-src']

                    # Get movie detail.
                    movie_detail_url = div_release_info.find('div', class_='release_text').span['data-url']
                    res_intro = requests.get(movie_detail_url)
                    soup_intro = BeautifulSoup(res_intro.text, 'html.parser')
                    intro_div = soup_intro.find('div', class_='movie_intro_info_r')
                    run_time = intro_div.find('span', string=re.compile('片　　長：')).text.replace('片　　長：', '').strip()
                    company = intro_div.find('span', string=re.compile('發行公司：')).text.replace('發行公司：', '').strip()

                    imdb_score_span = intro_div.find('span', string=re.compile('IMDb分數：'))
                    if imdb_score_span:
                        imdb_score = imdb_score_span.text.replace('IMDb分數：', '').strip()
                    else:
                        imdb_score = '無'

                    director_span = intro_div.find('span', class_='movie_intro_list')
                    child_tag = director_span.find('a')
                    if child_tag:
                        director = child_tag.text.strip()
                    else:
                        director = director_span.text.replace('導演：', '').strip()

                    cast_elements = director_span.find_next_sibling('span').find_all('a')
                    if len(cast_elements) > 0:
                        cast = ', '.join([cast_element.text.strip() for cast_element in cast_elements])
                    else:
                        cast_string = director_span.find_next_sibling('span').text.replace('演員：', '')
                        cast = ', '.join([c.strip() for c in cast_string.split('、')])

                    movie = {
                        '電影ID': movie_id,
                        '電影名稱': chinese_name,
                        '英文名稱': english_name,
                        '上映日期': release_time,
                        '片長': run_time,
                        '期待度': expectation,
                        '滿意度': satisfaction,
                        'IMDb分數': imdb_score,
                        '發行公司': company,
                        '導演': director,
                        '演員': cast,
                        '劇情介紹': intro,
                        '電影海報': img_url,
                    }
                    self.movies = pd.concat([self.movies, pd.DataFrame([movie])], ignore_index=True)
                    self.save_movies()

        self.movies.index += 1
        print('完成')

    def query_movie(self, this_week=False):
        if this_week:
            print('this_week')
            target_dates = []
            for i in range(7):
                target_dates.append(str(datetime.date.today() - datetime.timedelta(days=i)))

            print(target_dates)
            filt = self.movies['上映日期'].isin(target_dates)
            print('OK')
            result_df = self.movies.loc[filt]
            print(result_df)
        else:
            result_df = self.movies

        return result_df

    def get_movie_time(self, movie_id, date):
        if movie_id not in self.show_times['電影ID'].unique() or date not in self.show_times['日期'].unique():
            url = f'https://movies.yahoo.com.tw/ajax/pc/get_schedule_by_movie?movie_id={movie_id}&date={date}&area_id='
            res = requests.get(url)
            html = res.json().get('view')
            soup = BeautifulSoup(html, 'html.parser')

            # Areas
            area_timebox_divs = soup.find_all('div', class_='area_timebox')
            for area_timebox_div in area_timebox_divs:
                # Theaters
                theater_uls = area_timebox_div.find_all('ul')
                for theater_ul in theater_uls:
                    theater_name = theater_ul['data-theater_name']
                    theater_info_url = theater_ul['data-theater_schedules']

                    # If the theater is unknown, get theater info.
                    if theater_name not in self.theaters['戲院名稱'].unique():
                        print(f'新增戲院資訊: {theater_name}')
                        self.get_theater_info(theater_name, theater_info_url)

                    # Get movie time.
                    taps_lis = theater_ul.find_all('li', class_='taps')
                    for taps_li in taps_lis:
                        taps = ', '.join([span.text for span in taps_li.find_all('span')])

                        time_li = taps_li.find_next_sibling('li', class_='time _c')
                        time_labels = time_li.div.find_all('label')
                        for time_label in time_labels:
                            show_time = {
                                '電影ID': movie_id,
                                '戲院ID': theater_info_url.split('id=')[1],
                                '類型': taps,
                                '日期': date,
                                '時間': time_label.text,
                            }
                            self.show_times = pd.concat([self.show_times, pd.DataFrame([show_time])], ignore_index=True)

        filt = (self.show_times['電影ID'] == movie_id) & (self.show_times['日期'] == date)
        return self.show_times.loc[filt]

    def get_theater_info(self, theater_name, theater_info_url):
        res = requests.get(theater_info_url)
        soup = BeautifulSoup(res.text, 'html.parser')

        theater_id = theater_info_url.split('id=')[1]

        # Get theater address.
        ul_element = soup.find('div', class_='theaterlist_area').ul
        address_element = ul_element.find('li', string=re.compile('地址：'))
        address = address_element.text.replace('地址：', '').strip()

        # Get theater phone.
        phone_element = ul_element.find('li', string=re.compile('電話：'))
        phone = phone_element.text.replace('電話：', '').strip()
        phone = phone.replace('(', '').replace(')', '-')

        # Find coordinate by address.
        lat, lng = geocoder.arcgis(address).latlng

        # Get area from address.
        area = address[0:2].replace('臺', '台')

        theater = {
            '戲院ID': theater_id,
            '戲院名稱': theater_name,
            '地區': area,
            '電話': phone,
            '地址': address,
            '緯度': lat,
            '經度': lng,
        }

        self.theaters = pd.concat([self.theaters, pd.DataFrame([theater])], ignore_index=True)
        self.save_theaters()


def main():
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    main_window.display_loading_page()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
