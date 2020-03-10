# -*- coding: utf-8 -*-
import os
import argparse
from CheckPresentation.Main import Main
from CheckPresentation.Print import Print

r'''
Examples:
print D:\Presentations\003.pptx D:\Presentations\003.txt txt write
print D:\Presentations\003.pptx D:\Presentations\003.csv csv rewrite -e windows-1251

images D:\Presentations\003.pptx powerpoint
images D:\Presentations\003.pptx simple
images D:\Presentations\003.pptx skeleton

simple D:\Presentations\003.pptx -j
simple D:\Presentations\003.pptx
'''


def print_to_file(arguments):
    main = Main(arguments.path_to_presentation)
    result = main.Analyze.analyze()
    if arguments.extension == "txt":
        p = Print(result, arguments.path_to_out, arguments.path_to_presentation, encoding=arguments.encoding)
        print(p.txt(mode=arguments.write_mode))
    elif arguments.extension == "csv":
        p = Print(result, arguments.path_to_out, arguments.path_to_presentation, encoding=arguments.encoding)
        print(p.csv(mode=arguments.write_mode))


def generate_images(arguments):
    def print_result(res):
        print("Пути к картинкам:")
        for r in res:
            print(r)

    main = Main(arguments.path_to_presentation)
    images = main.Images
    if arguments.type == "powerpoint":
        print_result(main.Testing.generate_slide_images())
    elif arguments.type == "simple":
        print_result(images.generate())
    elif arguments.type == "skeleton":
        print_result(images.generate_skeleton())
    else:
        return "Укажите верный тип генерации."


def simple_result(arguments):
    main = Main(arguments.path_to_presentation)
    result = main.Analyze.analyze()
    if arguments.json:
        result = main.Utils.to_json(result)
        print(result)
        return
    for r in result:
        print(f"{r} => {result[r]}")


def create_parser():
    description = (
        '© newfox79 https://github.com/newfox79/ \n'
        'Приложение создано в помощь ученикам и экспертам в проверке задания 13.1 ОГЭ 2020. \n'
        'Приложение предоставляется "As Is" без каких либо дополнительных гарантий, и всё ещё находится в стадии '
        'активной разработки.\n'
        'По любым вопросам или предложениям можно написать в почту lis@chaikovskie.com'
    )
    parser = argparse.ArgumentParser(description=description)
    sub_parser = parser.add_subparsers()

    parser_print = sub_parser.add_parser("print",
                                         help="Укажите путь к презентации, путь к файлу в который необходимо выгрузить "
                                              "результаты, его расширение, режим записи. Примеры:\n"
                                              "ФАЙЛ_ПРЕЗЕНАТЦИИ -p C:/path/out.csv csv write | "
                                              "ФАЙЛ_ПРЕЗЕНТАЦИИ -p C:/path/out.txt txt rewrite")
    parser_print.add_argument("path_to_presentation", type=str, help="Путь к презентации формата ABSPATH")
    parser_print.add_argument("path_to_out", type=str, help="Путь к файлу выгрузки результатов формата ABSPATH")
    parser_print.add_argument("extension", type=str, help="Расширение файла выгрузки", choices=["txt", "csv"])
    parser_print.add_argument("write_mode", type=str, help="Режим записи файла", choices=["write", "rewrite"])
    parser_print.add_argument("-e", "--encoding", type=str, help="Кодировка файла выгрузки", default="utf-8")
    parser_print.set_defaults(func=print_to_file)

    parser_images = sub_parser.add_parser("images",
                                          help="Укажите путь к презентации, каким способом генерировать картинки, "
                                               "и вы получите абсолютные пути к сгенерированным картинкам")
    parser_images.add_argument("path_to_presentation", type=str, help="Путь к презентации формата ABSPATH")
    parser_images.add_argument("type", type=str, choices=["powerpoint", "simple", "skeleton"],
                               help="Способ генерации картинок. powerpoint - самый точный, но требует установленного "
                                    "powerpoint. simple - не точный, но быстрый, позволяет увидеть текст и картинки "
                                    "на презентации. skeleton - позволяет увидеть расположение элементов на презентации,"
                                    " синие прямоугольники: картинки, жёлтые: текст.")
    parser_images.set_defaults(func=generate_images)

    parser_simple = sub_parser.add_parser("simple",
                                          help="Укажите путь к презентации, и получите результат проверки")
    parser_simple.add_argument("path_to_presentation", type=str, help="Путь к презентации формата ABSPATH")
    parser_simple.add_argument("-j", "--json", action="store_true", help="Если нужен результат в json")
    parser_simple.set_defaults(func=simple_result)

    parser = parser.parse_args()
    return parser


if __name__ == '__main__':
    args = create_parser()
    args.func(args)
