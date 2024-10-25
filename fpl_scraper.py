import datetime
import time
import openpyxl
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def driver_setup():
    driver_path = r"C:\Users\Pc\Downloads\Programs\chromedriver-win64\chromedriver-win64\chromedriver.exe"
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service)
    return driver



def scrape_filters_data(driver, URL):
    driver.get(URL)
    driver.implicitly_wait(2)

    # Getting The Filtering Options:
    filters = driver.find_elements(By.XPATH,"//th[(@class='thead desc_only top sorting_asc_disabled'\
     or @class='thead desc_only sorting_asc_disabled') and @aria-controls='playerdata']")


    #Activate This If You Want To See The Filters & Their Index 👇:
    '''
     for idx, filter in enumerate(filters):
        clean_answer = filter.text.replace('\n',' ').strip()
        print(f"{idx} : {clean_answer}")
        
     Result :-
     0:Points Per Game/ 1:Last GW Points/ 2:Form/ 3:Value Form/ 4:Expected Points Next/ 5:ICT Index/ 6:Bonus System 
     '''

    # Getting Max & Min Value For Each Filter To Create An Overall Score Based On The 7 Filters:
    wait = WebDriverWait(driver,3)
    filters_min_and_max_scores = dict()

    for filter_idx in range(7):
        filters[filter_idx].click()

        scores_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td[@class= 'sorting_1']")))
        scores = tuple(float(score.text) for score in scores_elements)

        max_score = max(scores)
        min_score = min(scores)

        filters_min_and_max_scores[filters[filter_idx].text.replace("\n"," ").strip()] = (min_score, max_score)

        # Activate This To See MAX & MIN For Each Filter 👇:
        '''
        for key, value in filters_min_and_max_scores.items():
            print(f"{key}: {value}")
        '''
    return filters_min_and_max_scores
    driver.quit()



def scrape_players_scores_data(driver, URL):
    driver.get(URL)
    players_scores = dict()
    wait = WebDriverWait(driver,3)

    # This Will Get The Player Row:
    players_elements = tuple(wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH,"//tr[contains(@class,'odd') or contains(@class,'even')]"))))

    # This Will Get The Data In The Filters We Want:
    for element in players_elements:
        scores = element.find_elements(By.TAG_NAME,"td")
        scores = scores[0:1] + scores[5:12]
        scores_values = tuple(score.text for score in scores)

        players_scores[scores_values[0]] = scores_values[1:]

    return players_scores
    driver.quit()



def get_players_final_score(filters_data, players_scores_data):
    players_final_score = dict()
    filters = ["Points Per Game","Event Points","Form","Value Form","EP Next","ICT Index","BPS"]

    for player_name, scores in players_scores_data.items():
        final_score = 0
        for idx, score in enumerate(scores):

            # Creating a Score for 0 To 1 For Each Filter: (Score - Min) / (Max - Min)
            value_for_filter = (float(score) - filters_data.get(filters[idx])[0] )\
                               / ( filters_data.get(filters[idx])[1] - filters_data.get(filters[idx])[0] )
            final_score += value_for_filter
        players_final_score[player_name] = round(final_score,2)

    # Sorting The Dictionary Based On Values:
    players_final_score_sorted = dict(sorted(players_final_score.items(), key= lambda item: item[1], reverse=True))

    return players_final_score_sorted



def exporting_to_Excel(players_final_score):
    while True:
        choice = input("Do You Want To Export Players' Data Into Excel ? (Y/N):")

        if choice.lower() == "y":
            workbook = openpyxl.Workbook()
            sheet = workbook.active

            # Adding a Header:
            sheet.append(["Player","Final Score"])

            # Writing Data To The Sheet:
            for player, final_score in players_final_score.items():
                sheet.append([player,final_score])

            # Saving The Data:
            file_name = f"./Players Data/{datetime.date.today()}.xlsx"
            workbook.save(file_name)

            print("\nData Saved Successfully !!")
            print("Here Is The Top 25 Picks For Next Gameweek . . .")
            break

        elif choice.lower() == "n":
            print("No Problem, Here Is The Top 25 Picks For Next Gameweek . . .")
            break
        else:
            print("Please Enter A Valid Option !!\n")



def plotting_result(players_final_scores):
    # Adding Data Labels To The Plot:
    plt.figure(figsize=(14,7))
    bars = plt.bar(players_final_scores.keys(),players_final_scores.values())
    for bar in bars:
        y_value = bar.get_height()
        plt.text(x=bar.get_x() + bar.get_width() / 2,y=y_value,s=y_value,ha="center",va="bottom")

        if y_value >= 5:
            bar.set_color("purple")
        elif 4.5 <= y_value < 5:
            bar.set_color("navy")
        elif 4 <= y_value < 4.5:
            bar.set_color("green")
        else:
            bar.set_color("orange")

    plt.xticks(rotation=40)
    plt.title("Best 25 Picks For Next Gameweek")
    plt.yticks(range(0,8))
    plt.grid(axis="y")
    plt.show()



def main():
    URL = r"https://www.fplform.com/fpl-player-data"
    driver = driver_setup()

    filters_data= scrape_filters_data(driver, URL)
    players_scores_data= scrape_players_scores_data(driver, URL)

    players_final_score= get_players_final_score(filters_data,players_scores_data)
    top25_players_final_score= dict(list(get_players_final_score(filters_data,players_scores_data).items())[:25])

    exporting_to_Excel(players_final_score)
    plotting_result(top25_players_final_score)



if __name__ == "__main__":
    main()