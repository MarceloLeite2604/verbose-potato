from configuration import retrieve_output_file_path
import json
from collections import defaultdict
from bs4 import BeautifulSoup
import itertools
from util.cell import split_cell_reference
import htmlmin


def _append_style_tag(soup: BeautifulSoup, css_classes: list):
    style_tag = soup.new_tag('style')

    style_tag.string = '\n'.join(css_classes)

    if soup.head:
        soup.head.append(style_tag)
    else:
        head_tag = soup.new_tag("head")
        head_tag.append(style_tag)
        soup.html.insert(0, head_tag)

    return soup


def _create_css_classes(
        html_tags_by_style_properties_combinations: list,
        html_tag_attributes: defaultdict):

    css_classes = []
    class_count = 1

    total_html_tags_by_style_properties_entries = len(
        html_tags_by_style_properties_combinations)

    previous_percentage_displayed = 0

    for html_tags_by_style_properties_processed, (style_properties, html_tags) \
            in enumerate(html_tags_by_style_properties_combinations):

        if (percentage_displayed := int(html_tags_by_style_properties_processed / total_html_tags_by_style_properties_entries * 100)) \
                != previous_percentage_displayed:
            print(
                f'{percentage_displayed}% of HTML tags by style properties processed.')
            previous_percentage_displayed = percentage_displayed

        html_tags_to_associate_class = []

        create_css_class = False
        for html_tag in html_tags:
            html_tag_remaining_style_properties = html_tag_attributes[html_tag]['style']

            resolved_style_properties = html_tag_remaining_style_properties.intersection(
                set(style_properties))

            if resolved_style_properties:
                create_css_class = True

                html_tag_attributes[html_tag]['style'] = html_tag_remaining_style_properties - \
                    resolved_style_properties

                html_tags_to_associate_class.append(html_tag)

        if not create_css_class:
            continue

        class_name = f'style{class_count}'
        css_classes.append(
            f'.{class_name} {{\n    {"; ".join(style_properties)}\n }}')
        class_count += 1

        for html_tag_to_associate_class in html_tags_to_associate_class:
            html_tag_attributes[html_tag_to_associate_class]['class'].add(
                class_name)

    for html_tag, attributes in html_tag_attributes.items():

        if len(attributes['class']) > 0:
            html_tag['class'] = ' '.join(attributes['class'])
        else:
            if 'class' in html_tag.attrs:
                del html_tag['class']

        if len(attributes['style']) > 0:
            html_tag['style'] = '; '.join(attributes['style'])
        else:
            if 'style' in html_tag.attrs:
                del html_tag['style']

    return css_classes, html_tag_attributes


def _elaborate_html_tags_by_style_properties_combinations(
        html_tags_by_style_properties_map: defaultdict):

    html_tags_by_style_properties_combinations = defaultdict(set)

    total_style_properties_groups = len(html_tags_by_style_properties_map)
    previous_percentage_displayed = 0

    for style_properties_groups_processed, style_properties in enumerate(html_tags_by_style_properties_map):

        if (percentage_displayed := int(style_properties_groups_processed / total_style_properties_groups * 100)) \
                != previous_percentage_displayed:
            print(f'{percentage_displayed}% of style properties groups processed.')
            previous_percentage_displayed = percentage_displayed

        for subset_size in range(1, len(style_properties) + 1):
            for style_properties_subset in itertools.combinations(style_properties, subset_size):
                html_tags = html_tags_by_style_properties_map[style_properties]
                html_tags_by_style_properties_combinations[style_properties_subset].update(
                    html_tags)

    # Remove groups that are not worth creating a CSS class for
    style_properties_to_remove = [style_properties
                                  for style_properties, html_tags
                                  in html_tags_by_style_properties_combinations.items()
                                  if len(html_tags) <= 1]

    for style_properties_to_remove in style_properties_to_remove:
        del html_tags_by_style_properties_combinations[style_properties_to_remove]

    # Groups which attend more HTML tags and contain more properties must be resolved first.
    html_tags_by_style_properties_combinations = sorted(
        html_tags_by_style_properties_combinations.items(),
        key=lambda tuple: len(tuple[0])*100000 + len(tuple[1]), reverse=True)

    return html_tags_by_style_properties_combinations


def _elaborate_html_tags_styles_structures(soup: BeautifulSoup):
    html_tags_by_style_properties_map = defaultdict(set)

    html_tag_attributes = defaultdict(
        lambda: {'class': set(), 'style': set()})

    html_tags_with_styles = soup.find_all(style=True)
    total_tags = len(html_tags_with_styles)
    previous_percentage_displayed = 0

    for tags_processed, html_tag_with_style in enumerate(html_tags_with_styles):
        style_properties = html_tag_with_style['style'].split(';')

        if (percentage_displayed := int(tags_processed / total_tags * 100)) \
                != previous_percentage_displayed:
            print(f'{percentage_displayed}% of HTML tags processed.')
            previous_percentage_displayed = percentage_displayed

        style_properties = tuple(
            sorted([style_property.strip()
                    for style_property
                    in style_properties
                    if style_properties]))

        html_tags_by_style_properties_map[style_properties].add(
            html_tag_with_style)

        html_tag_attributes[html_tag_with_style]['style'] = \
            set(style_properties)

    return html_tags_by_style_properties_map, html_tag_attributes


def _retrieve_soup_structure(html_file_path: str):

    with open(html_file_path, 'r') as html_file:
        html = html_file.read()

    return BeautifulSoup(html, 'html.parser')


def _optimize_html(html_file_path):
    structure_file_path = retrieve_output_file_path(
        'structure', html_file_path)

    soup = _retrieve_soup_structure(structure_file_path)

    # A map of which HTML tag contains which styles.

    html_tags_by_style_properties_map, html_tag_attributes = \
        _elaborate_html_tags_styles_structures(soup)

    html_tags_by_style_properties_combinations = \
        _elaborate_html_tags_by_style_properties_combinations(
            html_tags_by_style_properties_map)

    css_classes, html_tag_attributes = _create_css_classes(
        html_tags_by_style_properties_combinations, html_tag_attributes)

    soup = _append_style_tag(soup, css_classes)

    html_tags_with_id = soup.find_all(id=True)
    for html_tag_with_id in html_tags_with_id:

        cell_reference = split_cell_reference(html_tag_with_id['id'])

        if cell_reference:
            html_tag_with_id['id'] = \
                f'{cell_reference["start"]["col"]}{cell_reference["start"]["row"]}'

    optimized_html = str(soup)

    minified_html = htmlmin.minify(optimized_html)

    # Save the optimized HTML back to the file
    optimized_structure_file_path = retrieve_output_file_path(
        'structure', 'optimized', html_file_path)

    with open(optimized_structure_file_path, 'w') as optimized_structure_file:
        optimized_structure_file.write(minified_html)


def optimize_workbook_structure():
    metadata_file_path = retrieve_output_file_path(
        'structure', 'metadata.json')

    with open(metadata_file_path, 'r') as metadata_file:
        metadata = json.load(metadata_file)

    for worksheet_name, structure_file_name in metadata.items():

        print(f'Optimizing worksheet \"{worksheet_name}\" structure.')
        _optimize_html(structure_file_name)
