import csv
import errno
import os
import shutil

import numpy as np
from pywinauto.application import Application

locations = ['Loc_A', 'Loc_B', 'Loc_C']


def run(path='.'):
    app = start_app()

    empty_folder(path, 'results')
    safe_mkdir(os.path.join(path, 'tmp'))
    safe_mkdir(os.path.join(path, 'results'))

    samples = get_sample_names(path)
    for sample in samples:
        for location in locations:
            measure_sample(app, path, sample, location)
    summarize_results(path)

    close_app(app)


def measure_sample(app, path, sample, location):
    if move_files_to_tmp(path, sample, location):
        dlg = get_main_window(app)
        open_recipe_window(dlg)

        select_recipe(app, path, location)
        select_sample(app)
        save_result(app, path, sample, location)

        close_recipe_window(app)


def start_app():
    return Application(backend='uia').start('PATH TO EXE')


def move_files_to_tmp(path, sample, location):
    empty_folder(path, 'tmp')
    tmp_path = os.path.join(path, 'tmp')
    folder_path = os.path.join(path, 'samples', sample, location)
    try:
        for content in os.listdir(folder_path):
            content_path = os.path.join(folder_path, content)
            try:
                if os.path.isfile(content_path):
                    shutil.copy2(content_path, tmp_path)
            except Exception as e:
                print(e)
        condFolder = get_folders('_cnd', folder_path)
        for cond in condFolder:
            foldername = cond.split(os.sep)
            shutil.copytree(cond, os.path.join(tmp_path, foldername[-1]))
        if len(os.listdir(tmp_path)) == 0:
            return False
        else:
            return True
    except FileNotFoundError as not_found:
        print('Skipped:', not_found.filename)


def empty_folder(path, folder):
    folder = os.path.join(path, folder)
    for content in os.listdir(folder):
        content_path = os.path.join(folder, content)
        try:
            if os.path.isfile(content_path):
                os.unlink(content_path)
            elif os.path.isdir(content_path):
                shutil.rmtree(content_path)
        except Exception as e:
            print(e)


def get_main_window(app):
    dlg = app.window(best_match='MAIN WINDOW NAME')
    dlg.wait('visible')
    return dlg


def open_recipe_window(main_window):
    menu_bar = main_window.child_window(title="Application", auto_id="MenuBar", control_type="MenuBar")
    menu_bar.Recipe.click_input()
    main_window.menu_select('Recipe->Load Recipe')


def select_recipe(app, path, location):
    recipes = os.listdir(os.path.join(path, 'recipes'))
    recipe_name = [r for r in recipes if location in r][0]
    recipe_path = os.path.abspath(os.path.join(path, 'recipes', recipe_name))
    file_dlg = app.window(title='File Open')
    file_dlg.wait('visible')
    file_dlg.ComboBox.Edit.set_text(recipe_path)
    file_dlg.OpenButton3.click()


def select_sample(app):
    app.Window_(best_match='Dialog', top_level_only=True).child_window(best_match='Recipe').child_window(
        title="Execute...", auto_id="7", control_type="Button").click()
    fld_dlg = app.Window_(best_match='Dialog', top_level_only=True).child_window(best_match='Folder Select')
    fld_dlg.SelectButton.click()


def close_recipe_window(app):
    app.Window_(best_match='Dialog', top_level_only=True).child_window(best_match='Recipe').CloseButton2.click()


def close_app(app):
    dlg = get_main_window(app)
    dlg.TitleBar.CloseButton.click()


def save_result(app, path, sample, location):
    result_path = os.path.abspath(os.path.join(path, 'results', '{}_{}.rrf'.format(sample, location)))
    result_dlg = app.Window_(best_match='Dialog', top_level_only=True).child_window(best_match='Recipe Result')
    result_dlg.Pane1.Button2.click()
    file_dlg = app.Window_(best_match='Dialog', top_level_only=True)
    file_dlg.ComboBox.Edit.set_text(result_path)
    file_dlg.SaveButton.click()
    result_dlg.CloseButton2.click()


def get_folders(ending, path='data'):
    folderOnly = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        for dirname in dirnames:
            if dirname.endswith(ending):
                folderOnly.append(os.sep.join([dirpath, dirname]))
    return folderOnly


def get_files(extension, path='data'):
    fileOnly = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        for filename in filenames:
            if filename.endswith(extension):
                fileOnly.append(os.sep.join([dirpath, filename]))
    return fileOnly


def get_sample_names(path):
    sample_path = os.path.join(path, 'samples')
    return [f for f in os.listdir(sample_path) if os.path.isdir(os.path.join(sample_path, f))]


def safe_mkdir(dir):
    try:
        os.makedirs(dir)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise


def summarize_sample_results(path, sample):
    data = {}
    result_path = os.path.join(path, 'results')
    result_files = sorted(os.listdir(result_path))
    result_files = [f for f in result_files if sample in f]
    for file in result_files:
        with open(os.path.join(result_path, file), 'r') as f:
            _, location = file[:-4].split('_')
            score = []
            width = []
            property_A = []
            property_B = []
            measurement_flag = False
            for line in f.readlines():
                content = line.split('\t')
                if content[0] == 'No.' and content[1] == 'Image' and 'Property A':
                    measurement_flag = True
                elif measurement_flag and '.tif' in content[1]:
                    if content[3].replace('.', '', 1).isdigit():
                        score.append(float(content[3]))
                    if content[4].replace('.', '', 1).isdigit():
                        width.append(float(content[4]))
                    if content[5].replace('.', '', 1).isdigit():
                        property_A.append(float(content[5]))
                    if content[6].replace('.', '', 1).isdigit():
                        property_B.append(float(content[6]))

            data[location] = {'score': np.mean(score),
                              'width': np.mean(width),
                              'property_A': np.mean(property_A),
                              'property_B': np.mean(property_B)}
    write_results(path, data, sample)


def write_results(path, data, sample):
    result_path = os.path.join(path, 'results')
    locations_measured = sorted(data.keys())
    with open(os.path.join(result_path, 'summary.csv'), 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)

        writer.writerow([sample])

        header1 = []
        header2 = []
        measurements = []
        for location in locations_measured:
            header1.append(location)
            for i in range(3):
                header1.append('')
            for field in ['score', 'width', 'property_A', 'property_B']:
                header2.append(field)
                measurements.append(data[location][field])

        writer.writerow(header1)
        writer.writerow(header2)
        writer.writerow(measurements)
        writer.writerow([])


def summarize_results(path):
    samples = get_sample_names(path)
    open(os.path.join(path, 'results', 'summary.csv'), 'w').close()
    for sample in samples:
        summarize_sample_results(path, sample)


if __name__ == '__main__':
    run()
