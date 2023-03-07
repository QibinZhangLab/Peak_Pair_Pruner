import PySimpleGUI as sG  # PySimpleGui website: https://www.pysimplegui.org/en/latest/
import xlsxwriter  # XlsxWriter website: https://pypi.org/project/XlsxWriter/
import os
from os.path import exists
from datetime import datetime


# convert_to_char() converts a column number into its corresponding .xlsx column letter code.
def convert_to_char(num):

    if num == 1:
        return 'A'
    elif num == 2:
        return 'B'
    elif num == 3:
        return 'C'
    elif num == 4:
        return 'D'
    elif num == 5:
        return 'E'
    elif num == 6:
        return 'F'
    elif num == 7:
        return 'G'
    elif num == 8:
        return 'H'
    elif num == 9:
        return 'I'
    elif num == 10:
        return 'J'
    elif num == 11:
        return 'K'
    elif num == 12:
        return 'L'
    elif num == 13:
        return 'M'
    elif num == 14:
        return 'N'
    elif num == 15:
        return 'O'
    elif num == 16:
        return 'P'
    elif num == 17:
        return 'Q'
    elif num == 18:
        return 'R'
    elif num == 19:
        return 'S'
    elif num == 20:
        return 'T'
    elif num == 21:
        return 'U'
    elif num == 22:
        return 'V'
    elif num == 23:
        return 'W'
    elif num == 24:
        return 'X'
    elif num == 25:
        return 'Y'
    elif num == 26:
        return 'Z'


# convert_to_grid() converts a row and column number into alphanumeric .xlsx grid code.
def convert_to_grid(row, column):
    row = row + 1
    column = column
    col_remainder_init = column % 26
    col_remainder = (column % 26) + 1
    col_cycles = int((column - col_remainder_init)/26)
    if col_cycles > 0:
        first_letter = convert_to_char(col_cycles)
    else:
        first_letter = ''
    second_letter = convert_to_char(col_remainder)
    return first_letter + second_letter + str(row)


# within_error_limit() checks that that the difference between two masses is within an acceptable error limit.
def within_error_limit(mass1, mass2, error):

    difference = float(mass1 - mass2)
    if difference < 0:
        difference = difference * - 1.0
    if difference <= error:
        outcome = True
    else:
        outcome = False
    return outcome


# selective_rejoin() is used in call_ppp() to selectively rebuild matrix rows.
def selective_rejoin(line_entrees, indexes, height_mat, average_mat, stdev_mat, end_line):

    new_line = ''
    past_first_word = False
    for index in indexes:
        if past_first_word:
            new_line += '\t'
        if index == height_mat or index == average_mat or index == stdev_mat:
            new_line += ' \t'
        new_line += line_entrees[index - 1]
        past_first_word = True
    if end_line:
        new_line += '\n'
    return new_line


# is_number() checks that a character is either a '.' or a number digit.
def is_number(character):

    if character == '.' or character == '0' or character == '1' or character == '2' or character == '3' or character == '4' or character == '5' or character == '6' or character == '7' or character == '8' or character == '9':
        return True
    else:
        return False


# str_is_number() checks that a string is a number.
def str_is_number(string):

    bad_character = False
    decimal_found = False
    for character in string:
        if not is_number(character):
            bad_character = True
        elif character == '.':
            if decimal_found:
                bad_character = True
            else:
                decimal_found = True
    if bad_character:
        return False
    else:
        return True


# str_is_number_pos_or_neg() checks that a user-provided string is a number (sign is irrelevant).
def str_is_number_pos_or_neg(string):

    bad_character = False
    count_char = - 1
    for character in string:
        count_char += 1
        if not is_number(character):
            if count_char == 0:
                if string[count_char] != '-':
                    bad_character = True
            else:
                bad_character = True
    if bad_character:
        return False
    else:
        return True


# str_is_integer() checks that a user-provided string is an integer number.
def str_is_integer(string):

    bad_character = False
    for character in string:
        if not is_number(character) or character == '.':
            bad_character = True
    if bad_character:
        return False
    else:
        return True


# mdf_check() checks that the user-provided mass defect window is valid.
def mdf_check(mass_defect, bot_of_range, top_of_range):
    if bot_of_range + top_of_range == 999 or bot_of_range + top_of_range == 1000:
        return True
    elif mass_defect >= bot_of_range:
        if mass_defect <= top_of_range:
            return True
        else:
            return False
    else:
        return False


# call_ppp() is the workhorse function launched by main_function() when the user clicks 'GO' in the GUI with valid entrees
def call_ppp(area_directory, mdf_bot, mdf_top, light_lth_r, heavy_htl_r, mix_lth_r, mix_lth_tol, tags, tag_shift, tag_tol, delete_multi, subtract_b, subtract_nh, processed_name, format_menu_choice, destination_directory, tag_string):

    # If bad_id_reporting is enabled, a line will print whenever a potential peak pair is pruned with a reason.
    # This feature is useful for code trouble-shooting.
    # Currently, bad_id_reporting can only be enabled by editing the boolean below.
    bad_id_reporting = False

    md_filter_range_bottom = mdf_bot
    md_filter_range_top = mdf_top
    minimum_light_light_to_heavy_ratio = light_lth_r
    minimum_heavy_heavy_to_light_ratio = heavy_htl_r
    theoret_mix_lth_ratio = mix_lth_r
    mix_lth_tolerance = mix_lth_tol
    tag_counts = tags
    tag_count_string = tag_string
    tag_light_to_heavy_shift = tag_shift
    identified_ppm_tolerance = unknown_ppm_tolerance = tag_light_to_heavy_tolerance = tag_tol
    isotope_mass_shifts = []
    mplusx_isotope_shifts = []
    exhaustive_tag_search = False
    if tag_counts[0] == 'all':
        exhaustive_tag_search = True
    else:
        for tag_count in tag_counts:
            isotope_mass_shift = tag_count * tag_light_to_heavy_shift
            isotope_mass_shifts += [isotope_mass_shift]

    delete_multimatches = delete_multi
    subtract_blank = subtract_b
    subtract_natural_heavy = subtract_nh
    format_choice = format_menu_choice
    output_file_name = processed_name
    if format_choice == 'Matrix':
        output_file_name += '.txt'
    elif format_choice == 'Report':
        output_file_name += '.xlsx'

    classes = []
    file_types = []
    injection_order = []
    batch_ids = []
    column_heads = []
    samples = []
    all_peaks = []

    light_indexes = []
    heavy_indexes = []
    mix_indexes = []
    qc_indexes = []
    sample_indexes = []
    blank_indexes = []

    height_matrix_start_pos = - 1
    average_matrix_start_pos = - 1
    stdev_matrix_start_pos = - 1
    multimatrix = False
    alignment_id_head_pos = - 1
    rt_head_pos = - 1
    mz_head_pos = - 1
    metabolite_head_pos = - 1
    adduct_head_pos = - 1
    formula_head_pos = - 1
    isotope_tracking_parent_id_head_pos = - 1
    isotope_tracking_weight_number_head_pos = - 1
    total_score_head_pos = - 1
    rt_similarity_head_pos = - 1
    dot_product_head_pos = - 1
    reverse_dot_product_head_pos = - 1
    s_to_n_head_pos = - 1
    msms_head_pos = - 1
    matched_ms2_head_pos = - 1
    manually_identified_head_pos = - 1

    classes_line = []
    file_types_line = []
    injection_order_line = []
    batch_id_statistics_line = []
    columns_line = []

    blank_index = 0
    light_index = 0
    heavy_index = 0
    mix_index = 0
    line_num = 0

    final_identified_peak_sets = []
    final_unknown_peak_sets = []
    final_multimatch_sets = []

    ppp_identified_sets = 0
    ppp_unknown_sets = 0
    msdial_identified_sets = 0
    msdial_unknown_sets = 0

    # The booleans below mostly relate to checking that the user has uploaded a correctly formatted MS-DIAL matrix file.
    # The related checks are not all-encompassing and may require expansion in the future.
    no_ppp = False
    no_class = True
    no_file_type = True
    no_heavy = True
    no_light = True
    no_mix = True
    no_blank = True
    no_injection_order = True
    no_batch_id = True
    no_alignment = True
    no_rt = True
    no_mz = True
    no_name = True
    for line in open(area_directory):
        line_num += 1
        line_entrees = line.split("\t")
        if line_num == 1:
            for entry in line_entrees:
                if entry == 'Class':
                    no_class = False
                elif entry[0:5] == 'Blank' or entry[0:5] == 'blank':
                    no_blank = False
                elif entry[0:5] == 'Light' or entry[0:5] == 'light':
                    no_light = False
                elif entry[0:5] == 'Heavy' or entry[0:5] == 'heavy':
                    no_heavy = False
                elif entry[0:3] == 'Mix' or entry[0:3] == 'mix':
                    no_mix = False
        elif line_num == 2:
            for entry in line_entrees:
                if entry == 'File type':
                    no_file_type = False
        elif line_num == 3:
            for entry in line_entrees:
                if entry == 'Injection order':
                    no_injection_order = False
        elif line_num == 4:
            for entry in line_entrees:
                if entry == 'Batch ID':
                    no_batch_id = False
        elif line_num == 5:
            for entry in line_entrees:
                if entry == 'Alignment ID':
                    no_alignment = False
                elif entry == 'Average Rt(min)':
                    no_rt = False
                elif entry == 'Average Mz':
                    no_mz = False
                elif entry == 'Metabolite name':
                    no_name = False
    if no_alignment or no_heavy or no_injection_order or no_batch_id or no_blank or no_class or no_file_type or no_light or no_mix or no_mz or no_name or no_rt:
        no_ppp = True
        print('Error: File to be analyzed is not an MS-DIAL Area or Height Matrix or is otherwise formatted incorrectly for this program.')
        return False

    # Here begins reading in the matrix file for processing.  The original matrix file is not edited in this process.
    line_num = 0
    if not no_ppp:
        for line in open(area_directory):
            line_num += 1
            line_entrees = line.split("\t")
            num_entrees = len(line_entrees)
            entry_count = 0
            if line_num == 1:
                hit_names = False
                hit_end = False
                classes_line = line_entrees
                found_blank = False
                found_light = False
                found_heavy = False
                found_mix = False
                for entry in line_entrees:
                    entry_count += 1
                    if entry == 'Class':
                        hit_names = True
                        height_matrix_start_pos = entry_count + 1
                    elif hit_names:
                        if entry[0:5] == 'Blank' or entry[0:5] == 'blank':
                            blank_indexes += [entry_count]
                            if not found_blank:
                                found_blank = True
                                blank_index = entry_count
                        elif entry[0:5] == 'Light' or entry[0:5] == 'light':
                            light_indexes += [entry_count]
                            if not found_light:
                                found_light = True
                                light_index = entry_count
                        elif entry[0:5] == 'Heavy' or entry[0:5] == 'heavy':
                            heavy_indexes += [entry_count]
                            if not found_heavy:
                                found_heavy = True
                                heavy_index = entry_count
                        elif entry[0:3] == 'Mix' or entry[0:3] == 'mix':
                            mix_indexes += [entry_count]
                            if not found_mix:
                                found_mix = True
                                mix_index = entry_count
                        elif entry_count != num_entrees and entry != 'NA' and not hit_end:
                            sample_indexes += [entry_count]
                        if entry_count == num_entrees:
                            length = len(entry)
                            entry = entry[:length-1]
                            hit_end = True
                        if entry == 'NA':
                            multimatrix = True
                            hit_end = True
                            if average_matrix_start_pos == - 1:
                                average_matrix_start_pos = entry_count
                        else:
                            classes += [entry]
            elif line_num == 2:
                hit_names = False
                file_types_line = line_entrees
                for entry in line_entrees:
                    entry_count += 1
                    file_types_line = line_entrees
                    if entry == 'File type':
                        hit_names = True
                    elif hit_names:
                        if entry_count == num_entrees:
                            length = len(entry)
                            entry = entry[:length-1]
                        if entry != 'NA':
                            file_types += [entry]
            elif line_num == 3:
                hit_names = False
                injection_order_line = line_entrees
                for entry in line_entrees:
                    entry_count += 1
                    if entry == 'Injection order':
                        hit_names = True
                    elif hit_names:
                        if entry_count == num_entrees:
                            length = len(entry)
                            entry = entry[:length-1]
                        if entry != 'NA':
                            injection_order += [entry]
            elif line_num == 4:
                hit_names = False
                batch_id_statistics_line = line_entrees
                hit_stdev = False
                for entry in line_entrees:
                    entry_count += 1
                    if entry == 'Batch ID':
                        hit_names = True
                    elif hit_names:
                        if entry_count == num_entrees:
                            length = len(entry)
                            entry = entry[:length-1]
                        if entry != 'NA' and entry != 'Average' and entry != 'Stdev':
                            batch_ids += [entry]
                        elif entry == 'Stdev' and not hit_stdev:
                            stdev_matrix_start_pos = entry_count
                            hit_stdev = True
            elif line_num == 5:
                hit_names = False
                columns_line = line_entrees
                found_blank = False
                found_light = False
                found_heavy = False
                found_mix = False
                for entry in line_entrees:
                    entry_count += 1
                    if entry == 'Alignment ID':
                        alignment_id_head_pos = entry_count
                    elif entry == 'Average Rt(min)':
                        rt_head_pos = entry_count
                    elif entry == 'Average Mz':
                        mz_head_pos = entry_count
                    elif entry == 'Metabolite name':
                        metabolite_head_pos = entry_count
                    elif entry == 'Adduct type':
                        adduct_head_pos = entry_count
                    elif entry == 'Formula':
                        formula_head_pos = entry_count
                    elif entry == 'MS/MS matched':
                        matched_ms2_head_pos = entry_count
                    elif entry == 'Manually modified for annotation':
                        manually_identified_head_pos = entry_count
                    elif entry == 'Isotope tracking parent ID':
                        isotope_tracking_parent_id_head_pos = entry_count
                    elif entry == 'Isotope tracking weight number':
                        isotope_tracking_weight_number_head_pos = entry_count
                    elif entry == 'Total score':
                        total_score_head_pos = entry_count
                    elif entry == 'RT similarity':
                        rt_similarity_head_pos = entry_count
                    elif entry == 'Dot product':
                        dot_product_head_pos = entry_count
                    elif entry == 'Reverse dot product':
                        reverse_dot_product_head_pos = entry_count
                    elif entry == 'S/N average':
                        s_to_n_head_pos = entry_count
                    elif entry == 'MS/MS spectrum':
                        hit_names = True
                        msms_head_pos = entry_count
                    elif hit_names and multimatrix and stdev_matrix_start_pos > entry_count >= average_matrix_start_pos:
                        if entry[0:5] == 'Blank' or entry[0:5] == 'blank':
                            blank_indexes += [entry_count]
                            if not found_blank:
                                found_blank = True
                                blank_index = entry_count
                        elif entry[0:5] == 'Light' or entry[0:5] == 'light':
                            light_indexes += [entry_count]
                            if not found_light:
                                found_light = True
                                light_index = entry_count
                        elif entry[0:5] == 'Heavy' or entry[0:5] == 'heavy':
                            heavy_indexes += [entry_count]
                            if not found_heavy:
                                found_heavy = True
                                heavy_index = entry_count
                        elif entry[0:3] == 'Mix' or entry[0:3] == 'mix':
                            mix_indexes += [entry_count]
                            if not found_mix:
                                found_mix = True
                                mix_index = entry_count
                        else:
                            sample_indexes += [entry_count]
                    if entry_count == num_entrees:
                        length = len(entry)
                        entry = entry[:length-1]
                    column_heads += [entry]
                    if hit_names and entry_count >= height_matrix_start_pos:
                        if average_matrix_start_pos > 0:
                            if entry_count < average_matrix_start_pos:
                                samples += [entry]
                        else:
                            samples += [entry]
            else:
                new_peak = []
                for entry in line_entrees:
                    entry_count += 1
                    if entry_count == num_entrees:
                        length = len(entry)
                        entry = entry[:length-1]
                    new_peak += [entry]
                all_peaks += [new_peak]

        # The original matrix file is now closed, and processing of the internalized data begins.
        # First, potential peak pairs are found based off isotopic relationships given by MS-DIAL and user tag parameters.
        position = -1
        if exhaustive_tag_search:
            tag_counts = []
            isotope_shifts_found = []
            print('')
            print('Executing preliminary search for possible tagging levels:')
            for peak1 in all_peaks:
                peak_id1 = peak1[alignment_id_head_pos-1]
                parent_id1 = peak1[isotope_tracking_parent_id_head_pos-1]
                if peak_id1 == parent_id1:
                    for peak2 in all_peaks:
                        parent_id2 = peak2[isotope_tracking_parent_id_head_pos-1]
                        if parent_id1 == parent_id2:
                            isotope1 = int(peak2[isotope_tracking_weight_number_head_pos-1])
                            isotope_already_found = False
                            if isotope1 > 0:
                                for isotope2 in isotope_shifts_found:
                                    if isotope1 == isotope2:
                                        isotope_already_found = True
                                if not isotope_already_found:
                                    isotope_shifts_found += [isotope1]
            common_denominator = int(round(tag_light_to_heavy_shift))
            for found_isotope in isotope_shifts_found:
                if float(found_isotope) % float(common_denominator) == 0.0:
                    new_count = int(found_isotope/common_denominator)
                    new_count_string = str(new_count)
                    tag_counts += [new_count]
                    print(new_count_string + ' tag(s) added as a possible tagging level.')
            tag_counts.sort()
            for tag_count in tag_counts:
                isotope_mass_shift = tag_count * tag_light_to_heavy_shift
                isotope_mass_shifts += [isotope_mass_shift]

        for tag in tag_counts:
            num_unidentified_matches = 0
            unknown_peak_sets = []
            num_identified_matches = 0
            identified_peak_sets = []
            num_multimatch_sets = 0
            multimatch_sets = []
            position += 1
            tag_count = tag_counts[position]
            isotope_mass_shift = isotope_mass_shifts[position]
            mplusx_isotope_shift = str(int(round(isotope_mass_shift)))
            print('')
            print('Processing peak pairs for ' + str(tag_count) + ' tags per molecule:')

            for peak1 in all_peaks:
                peak_id1 = peak1[alignment_id_head_pos-1]
                parent_id1 = peak1[isotope_tracking_parent_id_head_pos-1]
                if peak_id1 == parent_id1:
                    match_found = False
                    name = peak1[metabolite_head_pos-1]
                    known = False
                    unknown = False
                    if name == 'Unknown':
                        unknown = True
                    else:
                        known = True
                    new_set = []
                    new_set += [peak1]
                    for peak2 in all_peaks:
                        parent_id2 = peak2[isotope_tracking_parent_id_head_pos-1]
                        isotope = peak2[isotope_tracking_weight_number_head_pos-1]
                        if parent_id2 == peak_id1 and isotope == mplusx_isotope_shift:
                            match_found = True
                            new_set += [peak2]
                    if match_found:
                        if known:
                            identified_peak_sets += [new_set]
                            num_identified_matches += 1
                        elif unknown:
                            unknown_peak_sets += [new_set]
                            num_unidentified_matches += 1
                    elif known and bad_id_reporting:
                        print('Bad ID (no initial pair): ' + name)

            print('')
            print('Potential peak pairs from Alignment matrix:')
            print('Identified peak pairs: ' + str(len(identified_peak_sets)))
            print('Unknown peak pairs: ' + str(len(unknown_peak_sets)))

            # Multimatches are here detected and separated.
            replacement = []
            for peaks in identified_peak_sets:
                if len(peaks) > 2:
                    multimatch_sets += [peaks]
                    num_multimatch_sets += 1
                    num_identified_matches -= 1
                else:
                    replacement += [peaks]
            identified_peak_sets = replacement

            replacement = []
            for peaks in unknown_peak_sets:
                if len(peaks) > 2:
                    multimatch_sets += [peaks]
                    num_multimatch_sets += 1
                    num_unidentified_matches -= 1
                else:
                    replacement += [peaks]
            unknown_peak_sets = replacement

            # Some temporarily disabled code here related to multimatch processing.
            # print('')
            # print('Multimatches separated:')
            # print('Identified sets: ' + str(len(identified_peak_sets)))
            msdial_identified_sets += len(identified_peak_sets)
            # print('Unknown sets: ' + str(len(unknown_peak_sets)))
            msdial_unknown_sets += len(unknown_peak_sets)
            # print('Multimatch sets: ' + str(len(multiMatch_sets)))
            msdial_multimatch_sets = len(multimatch_sets)

            # Here begins mass defect filtering and peak pair mass shift checks.
            replacement = []
            for peak_set in identified_peak_sets:
                peak_mz1 = float(peak_set[0][mz_head_pos-1])
                peak_mz2 = float(peak_set[1][mz_head_pos-1])
                mass_diff = peak_mz2 - peak_mz1
                charge_level = int(round(isotope_mass_shift / mass_diff))
                peak_mz1_mass_defect_mda = int(round(((peak_mz1 * charge_level) % 1) * 1000))
                if peak_mz1_mass_defect_mda > 499:
                    peak_mz1_mass_defect_mda -= 1000
                peak_mz2_mass_defect_mda = int(round(((peak_mz2 * charge_level) % 1) * 1000))
                if peak_mz2_mass_defect_mda > 499:
                    peak_mz2_mass_defect_mda -= 1000
                theoretical_mz2 = peak_mz1 + isotope_mass_shift / charge_level
                identified_mass_tolerance = identified_ppm_tolerance * peak_mz1 * 10 ** (-6)
                pass_mass_check = within_error_limit(peak_mz2, theoretical_mz2, identified_mass_tolerance)
                pass_mdf_check1 = mdf_check(peak_mz1_mass_defect_mda, md_filter_range_bottom, md_filter_range_top)
                pass_mdf_check2 = mdf_check(peak_mz2_mass_defect_mda, md_filter_range_bottom, md_filter_range_top)
                if pass_mass_check and pass_mdf_check1 and pass_mdf_check2:
                    replacement += [peak_set]
                elif bad_id_reporting:
                    if pass_mass_check:
                        print('Bad ID (outside MDF window): ' + peak_set[0][metabolite_head_pos-1])
                    else:
                        print('Bad ID (failed mass shift check): ' + peak_set[0][metabolite_head_pos-1])
            identified_peak_sets = replacement

            replacement = []
            for peak_set in unknown_peak_sets:
                peak_mz1 = float(peak_set[0][mz_head_pos-1])
                peak_mz2 = float(peak_set[1][mz_head_pos-1])
                mass_diff = peak_mz2 - peak_mz1
                charge_level = int(round(isotope_mass_shift / mass_diff))
                peak_mz1_mass_defect_mda = int(round(((peak_mz1 * charge_level) % 1) * 1000))
                if peak_mz1_mass_defect_mda > 499:
                    peak_mz1_mass_defect_mda -= 1000
                peak_mz2_mass_defect_mda = int(round(((peak_mz2 * charge_level) % 1) * 1000))
                if peak_mz2_mass_defect_mda > 499:
                    peak_mz2_mass_defect_mda -= 1000
                theoretical_mz2 = peak_mz1 + isotope_mass_shift / charge_level
                identified_mass_tolerance = unknown_ppm_tolerance * peak_mz1 * 10 ** (-6)
                pass_mass_check = within_error_limit(peak_mz2, theoretical_mz2, identified_mass_tolerance)
                pass_mdf_check1 = mdf_check(peak_mz1_mass_defect_mda, md_filter_range_bottom, md_filter_range_top)
                pass_mdf_check2 = mdf_check(peak_mz2_mass_defect_mda, md_filter_range_bottom, md_filter_range_top)
                if pass_mass_check and pass_mdf_check1 and pass_mdf_check2:
                    replacement += [peak_set]
            unknown_peak_sets = replacement

            replacement = []
            if delete_multimatches:
                multimatch_sets = []
            # This is where multimatch peak clusters are reduced to a single peak pair if not deleted outright.
            # As we have not had opportunity to properly test this feature, multimatches are currently deleted if detected.
            # We have not encountered any multimatches from MS-DIAL at this time.
            for peak_set in multimatch_sets:
                new_set = []
                new_set += [peak_set[0]]
                peak_mz1 = float(peak_set[0][mz_head_pos-1])
                mass_errors = []
                for j in range(len(peak_set) - 1):
                    next_peak = peak_set[j]
                    peak_mz2 = float(next_peak[mz_head_pos-1])
                    mass_diff = peak_mz2 - peak_mz1
                    charge_level = int(round(isotope_mass_shift / mass_diff))
                    theoretical_mz2 = peak_mz1 + isotope_mass_shift / charge_level
                    mass_error = abs(theoretical_mz2 - mass_diff)
                    mass_errors += mass_error
                min_error = min(mass_errors)
                index = mass_errors.index(min_error)
                new_set += peak_set[index + 1]
                peak_mz2 = float(new_set[1][mz_head_pos-1])
                mass_diff = peak_mz2 - peak_mz1
                charge_level = int(round(isotope_mass_shift / mass_diff))
                peak_mz1_mass_defect_mda = int(round(((peak_mz1 * charge_level) % 1) * 1000))
                if peak_mz1_mass_defect_mda > 499:
                    peak_mz1_mass_defect_mda -= 1000
                peak_mz2_mass_defect_mda = int(round(((peak_mz2 * charge_level) % 1) * 1000))
                if peak_mz2_mass_defect_mda > 499:
                    peak_mz2_mass_defect_mda -= 1000
                theoretical_mz2 = peak_mz1 + isotope_mass_shift / charge_level
                unknown_mass_tolerance = unknown_ppm_tolerance * peak_mz1 * 10 ** (-6)
                pass_mass_check = within_error_limit(peak_mz2, theoretical_mz2, unknown_mass_tolerance)
                pass_mdf_check1 = mdf_check(peak_mz1_mass_defect_mda, md_filter_range_bottom, md_filter_range_top)
                pass_mdf_check2 = mdf_check(peak_mz2_mass_defect_mda, md_filter_range_bottom, md_filter_range_top)
                if pass_mass_check and pass_mdf_check1 and pass_mdf_check2:
                    replacement += [peak_set]
            multimatch_sets = replacement

            print('')
            print('Mass checks performed with provided tolerances: ')
            print('Mass defect filter lower limit: ' + str(md_filter_range_bottom) + ' mDa')
            print('Mass defect filter upper limit: ' + str(md_filter_range_top) + ' mDa')
            print('Number of tags per molecule: ' + str(tag_count))
            print('Exact mass shift between light and heavy tags: ' + str(tag_light_to_heavy_shift) + ' Da')
            print('Mass shift tolerance: ' + str(tag_light_to_heavy_tolerance) + ' ppm')
            # print('Delete all multimatches: ' + str(delete_multimatches))
            print('')
            print('Identified peak pairs: ' + str(len(identified_peak_sets)))
            print('Unknown peak pairs: ' + str(len(unknown_peak_sets)))
            # print('Multimatch sets reduced to best match (by mass shift error) or deleted: ' + str(len(multiMatch_sets)))

            # Here begins quantitative corrections for background peaks and isotopic overlaps as well as QC checks.
            replacement = []
            for peak_set in identified_peak_sets:
                if subtract_blank:
                    for index in sample_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in mix_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in light_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in heavy_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                light_light_area = float(peak_set[0][light_index-1]) + 0.1
                light_heavy_area = float(peak_set[1][light_index-1]) + 0.1
                light_lth_ratio = light_light_area / light_heavy_area
                if subtract_natural_heavy:
                    for index in sample_indexes:
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[0][index-1]) * light_lth_ratio ** (-1)
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in mix_indexes:
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[0][index-1]) * light_lth_ratio ** (-1)
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                light_light_area = float(peak_set[0][light_index-1]) + 0.1
                light_heavy_area = float(peak_set[1][light_index-1]) + 0.1
                light_lth_ratio = light_light_area / light_heavy_area
                heavy_light_area = float(peak_set[0][heavy_index-1]) + 0.1
                heavy_heavy_area = float(peak_set[1][heavy_index-1]) + 0.1
                heavy_htl_ratio = heavy_heavy_area / heavy_light_area
                mix_light_area = float(peak_set[0][mix_index-1]) + 0.1
                mix_heavy_area = float(peak_set[1][mix_index-1]) + 0.1
                mix_lth = mix_light_area / mix_heavy_area
                mix_ratio_diff = abs(theoret_mix_lth_ratio - mix_lth)
                if light_lth_ratio >= minimum_light_light_to_heavy_ratio and heavy_htl_ratio >= minimum_heavy_heavy_to_light_ratio and mix_ratio_diff <= mix_lth_tolerance:
                    replacement += [peak_set]
                elif bad_id_reporting:
                    print('Bad ID (1 or more failed QC checks): ' + peak_set[0][metabolite_head_pos-1])
            identified_peak_sets = replacement

            replacement = []
            for peak_set in unknown_peak_sets:
                if subtract_blank:
                    for index in sample_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in mix_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in light_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in heavy_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                light_light_area = float(peak_set[0][light_index-1]) + 0.1
                light_heavy_area = float(peak_set[1][light_index-1]) + 0.1
                light_lth_ratio = light_light_area / light_heavy_area
                if subtract_natural_heavy:
                    for index in sample_indexes:
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[0][index-1]) * light_lth_ratio ** (-1)
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in mix_indexes:
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[0][index-1]) * light_lth_ratio ** (-1)
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                light_light_area = float(peak_set[0][light_index-1]) + 0.1
                light_heavy_area = float(peak_set[1][light_index-1]) + 0.1
                light_lth_ratio = light_light_area / light_heavy_area
                heavy_light_area = float(peak_set[0][heavy_index-1]) + 0.1
                heavy_heavy_area = float(peak_set[1][heavy_index-1]) + 0.1
                heavy_htl_ratio = heavy_heavy_area / heavy_light_area
                mix_light_area = float(peak_set[0][mix_index-1]) + 0.1
                mix_heavy_area = float(peak_set[1][mix_index-1]) + 0.1
                mix_lth = mix_light_area / mix_heavy_area
                mix_ratio_diff = abs(theoret_mix_lth_ratio - mix_lth)
                if light_lth_ratio >= minimum_light_light_to_heavy_ratio and heavy_htl_ratio >= minimum_heavy_heavy_to_light_ratio and mix_ratio_diff <= mix_lth_tolerance:
                    replacement += [peak_set]
            unknown_peak_sets = replacement

            replacement = []
            for peak_set in multimatch_sets:
                if subtract_blank:
                    for index in sample_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in mix_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in light_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in heavy_indexes:
                        temp_val = float(peak_set[0][index-1]) - float(peak_set[0][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[0][index-1] = str(temp_val)
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[1][blank_index-1])
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                light_light_area = float(peak_set[0][light_index-1]) + 0.1
                light_heavy_area = float(peak_set[1][light_index-1]) + 0.1
                light_lth_ratio = light_light_area / light_heavy_area
                if subtract_natural_heavy:
                    for index in sample_indexes:
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[0][index-1]) * light_lth_ratio ** (-1)
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                    for index in mix_indexes:
                        temp_val = float(peak_set[1][index-1]) - float(peak_set[0][index-1]) * light_lth_ratio ** (-1)
                        if temp_val < 0:
                            temp_val = 0
                        peak_set[1][index-1] = str(temp_val)
                light_light_area = float(peak_set[0][light_index-1]) + 0.1
                light_heavy_area = float(peak_set[1][light_index-1]) + 0.1
                light_lth_ratio = light_light_area / light_heavy_area
                heavy_light_area = float(peak_set[0][heavy_index-1]) + 0.1
                heavy_heavy_area = float(peak_set[1][heavy_index-1]) + 0.1
                heavy_htl_ratio = heavy_heavy_area / heavy_light_area
                mix_light_area = float(peak_set[0][mix_index-1]) + 0.1
                mix_heavy_area = float(peak_set[1][mix_index-1]) + 0.1
                mix_lth = mix_light_area / mix_heavy_area
                mix_ratio_diff = abs(theoret_mix_lth_ratio - mix_lth)
                if light_lth_ratio >= minimum_light_light_to_heavy_ratio and heavy_htl_ratio >= minimum_heavy_heavy_to_light_ratio and mix_ratio_diff <= mix_lth_tolerance:
                    replacement += [peak_set]
            multimatch_sets = replacement

            print('')
            if subtract_blank:
                print('Background peak values subtracted from samples and QCs.')
            if subtract_natural_heavy:
                print('Isotopic overlap corrections applied to samples and mix QCs.')
            print('')
            print('Isotopic ratio checks performed with provided parameters: ')
            print('Minimum light QC L/H ratio: ' + str(minimum_light_light_to_heavy_ratio))
            print('Minimum heavy QC H/L ratio: ' + str(minimum_heavy_heavy_to_light_ratio))
            print('Mix QC theoretical L/H ratio: ' + str(theoret_mix_lth_ratio))
            print('Mix QC ratio tolerance: ' + str(mix_lth_tolerance))
            print('')
            print('Identified peak pairs: ' + str(len(identified_peak_sets)))
            ppp_identified_sets += len(identified_peak_sets)
            print('Unknown peak pairs: ' + str(len(unknown_peak_sets)))
            ppp_unknown_sets += len(unknown_peak_sets)
            # print('Multimatch sets: ' + str(len(multiMatch_sets)))
            # ppp_multimatch_sets = len(multimatch_sets)

            for peak_set in identified_peak_sets:
                if format_choice == 'Report':
                    peak_set += [int(tag_count)]
                final_identified_peak_sets += [peak_set]
            for peak_set in unknown_peak_sets:
                if format_choice == 'Report':
                    peak_set += [int(tag_count)]
                final_unknown_peak_sets += [peak_set]

        # Now begins outputting to a new file.
        name_and_directory = os.path.join(destination_directory, output_file_name)

        if format_choice == 'Matrix':

            with open(name_and_directory, 'a') as f:

                indexes = [alignment_id_head_pos, rt_head_pos, mz_head_pos, metabolite_head_pos, adduct_head_pos, formula_head_pos, msms_head_pos, height_matrix_start_pos]
                num_entrees = len(column_heads)
                remaining = num_entrees - height_matrix_start_pos
                for k in range(remaining):
                    indexes += [k + height_matrix_start_pos + 1]

                end_line = True

                line = selective_rejoin(classes_line, indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                f.write(line)
                line = selective_rejoin(file_types_line, indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                f.write(line)
                line = selective_rejoin(injection_order_line, indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                f.write(line)
                line = selective_rejoin(batch_id_statistics_line, indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                f.write(line)
                line = selective_rejoin(columns_line, indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                f.write(line)
                for peak_set in final_identified_peak_sets:
                    line = selective_rejoin(peak_set[0], indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                    f.write(line)
                    line = selective_rejoin(peak_set[1], indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                    f.write(line)
                    f.write('\n')
                for peak_set in final_unknown_peak_sets:
                    line = selective_rejoin(peak_set[0], indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                    f.write(line)
                    line = selective_rejoin(peak_set[1], indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                    f.write(line)
                    f.write('\n')
                for peak_set in final_multimatch_sets:
                    line = selective_rejoin(peak_set[0], indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                    f.write(line)
                    line = selective_rejoin(peak_set[1], indexes, height_matrix_start_pos, average_matrix_start_pos, stdev_matrix_start_pos, end_line)
                    f.write(line)
                    f.write('\n')

                f.write('END\n')
                f.write('PPP v1.1 modified alignment matrix with the following PPP parameters:' + '\n')
                f.write('Alignment matrix analyzed: ' + str(area_directory) + '\n')
                f.write('Mass defect filter lower limit in mDa: ' + '\t' + str(md_filter_range_bottom) + '\n')
                f.write('Mass defect filter upper limit in mDa:' + '\t' + str(md_filter_range_top) + '\n')
                f.write('Minimum light QC L/H ratio: ' + '\t' + str(minimum_light_light_to_heavy_ratio) + '\n')
                f.write('Minimum heavy QC H/L ratio: ' + '\t' + str(minimum_heavy_heavy_to_light_ratio) + '\n')
                f.write('Mix QC theoretical L/H ratio: ' + '\t' + str(theoret_mix_lth_ratio) + '\n')
                f.write('Mix QC ratio tolerance: ' + '\t' + str(mix_lth_tolerance) + '\n')
                f.write('Number of tags per molecule: ' + '\t' + str(tag_count) + '\n')
                f.write('Exact mass shift between light and heavy tags in Da: ' + '\t' + str(tag_light_to_heavy_shift) + '\n')
                f.write('Mass shift tolerance in ppm: ' + '\t' + str(tag_light_to_heavy_tolerance) + '\n')
                # f.write('Delete multimatches?: ' + '\t' + str(delete_multimatches) + '\n')
                f.write('Subtract background values from samples and QCs?: ' + '\t' + str(subtract_blank) + '\n')
                f.write('Subtract isotopic overlap from samples and Mix QCs?: ' + '\t' + str(subtract_natural_heavy) + '\n')

        elif format_choice == 'Report':

            workbook = xlsxwriter.Workbook(name_and_directory)
            worksheet = workbook.add_worksheet("PPP Results")

            final_sample_signal_indexes = []
            final_sample_average_indexes = []
            final_sample_stdev_indexes = []
            final_mix_signal_indexes = []
            final_mix_average_indexes = []
            final_mix_stdev_indexes = []
            if multimatrix:
                for index in sample_indexes:
                    if index < average_matrix_start_pos:
                        final_sample_signal_indexes += [index]
                    elif index < stdev_matrix_start_pos:
                        final_sample_average_indexes += [index]
                    else:
                        final_sample_stdev_indexes += [index]
                for index in mix_indexes:
                    if index < average_matrix_start_pos:
                        final_mix_signal_indexes += [index]
                    elif index < stdev_matrix_start_pos:
                        final_mix_average_indexes += [index]
                    else:
                        final_mix_stdev_indexes += [index]
            else:
                final_sample_signal_indexes = sample_indexes
                final_mix_signal_indexes = mix_indexes

            peakpair_columns = ['Metabolite', 'Alignment ID', 'Average RT (min)', 'Average MZ (Light)', 'Average MZ (Heavy)', 'Num Tags', 'Adduct', 'ID type', 'Total Score']
            num_peakpair_columns = len(peakpair_columns)
            num_sample_average_columns = len(final_sample_average_indexes)
            num_sample_signal_columns = len(final_sample_signal_indexes)
            num_mix_average_columns = len(final_mix_average_indexes)
            num_mix_signal_columns = len(final_mix_signal_indexes)
            num_data_columns = num_sample_average_columns + num_sample_signal_columns + num_mix_average_columns + num_mix_signal_columns
            num_total_columns = num_peakpair_columns + num_data_columns

            merge_format1 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#CCECFF'
                                                })
            merge_format2 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#99FF99'
                                                })
            merge_format3 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#FF9999'
                                                })
            merge_format4 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#CCFFCC'
                                                })
            merge_format5 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#FFFFCC'
                                                })
            merge_format6 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#FFCCCC'
                                                })
            merge_format7 = workbook.add_format({
                                                'bold':     True,
                                                'border':   5,
                                                'align':    'center',
                                                'valign':   'vcenter',
                                                'fg_color': '#FFCCFF'
                                                })

            text_format = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#CCECFF'
                                                })
            number_format2 = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#99FF99'
                                                })
            number_format2.set_num_format('0.000')
            number_format3 = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#FF9999'
                                                })
            number_format3.set_num_format('0.000')
            number_format4 = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#CCFFCC'
                                                })
            number_format4.set_num_format('0.000')
            number_format5 = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#FFFFCC'
                                                })
            number_format5.set_num_format('0.000')
            number_format6 = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#FFCCCC'
                                                })
            number_format6.set_num_format('0.000')
            number_format7 = workbook.add_format({
                                                'bold':     False,
                                                #'border':   5,
                                                'align':    'left',
                                                'valign':   'vcenter',
                                                'fg_color': '#FFCCFF'
                                                })
            number_format7.set_num_format('0.000')
            row = 0
            worksheet.write_row(row, 0, ['PPP v1.1 report with the following PPP parameters:'])
            row += 1
            row += 1
            worksheet.write_row(row, 0, ['Alignment matrix analyzed: ' + area_directory])
            row += 1
            worksheet.write_row(row, 0, ['Mass defect filter lower limit in mDa: ' + str(md_filter_range_bottom)])
            row += 1
            worksheet.write_row(row, 0, ['Mass defect filter upper limit in mDa: ' + str(md_filter_range_top)])
            row += 1
            worksheet.write_row(row, 0, ['Minimum light QC L/H ratio: ' + str(minimum_light_light_to_heavy_ratio)])
            row += 1
            worksheet.write_row(row, 0, ['Minimum heavy QC H/L ratio: ' + str(minimum_heavy_heavy_to_light_ratio)])
            row += 1
            worksheet.write_row(row, 0, ['Mix QC theoretical L/H ratio: ' + str(theoret_mix_lth_ratio)])
            row += 1
            worksheet.write_row(row, 0, ['Mix QC ratio tolerance: ' + str(mix_lth_tolerance)])
            row += 1
            worksheet.write_row(row, 0, ['Number of tags per molecule: ' + str(tag_counts)])
            row += 1
            worksheet.write_row(row, 0, ['Exact mass shift between light and heavy tags in Da: ' + str(tag_light_to_heavy_shift)])
            row += 1
            worksheet.write_row(row, 0, ['Mass shift tolerance in ppm: ' + str(tag_light_to_heavy_tolerance)])
            row += 1
            # worksheet.write_row(row, 0, ['Delete multimatches?: ' + str(delete_multimatches)])
            # row += 1
            worksheet.write_row(row, 0, ['Subtract background values from samples and QCs?: ' + str(subtract_blank)])
            row += 1
            worksheet.write_row(row, 0, ['Subtract isotopic overlap from samples and mix QCs?: ' + str(subtract_natural_heavy)])
            row += 1
            row += 1
            worksheet.write_row(row, 0, ['Potential peak pairs imported from MS-DIAL:'])
            row += 1
            row += 1
            worksheet.write_row(row, 0, ['Identified: ' + str(msdial_identified_sets)])
            row += 1
            worksheet.write_row(row, 0, ['Unknown: ' + str(msdial_unknown_sets)])
            row += 1
            # worksheet.write_row(row, 0, ['Multimatch: ' + str(msdial_multimatch_sets)])
            # row += 1
            row += 1
            worksheet.write_row(row, 0, ['PPP validated peak pairs:'])
            row += 1
            row += 1
            worksheet.write_row(row, 0, ['Identified: ' + str(ppp_identified_sets)])
            row += 1
            worksheet.write_row(row, 0, ['Unknown: ' + str(ppp_unknown_sets)])
            row += 1
            # worksheet.write_row(row, 0, ['Multimatch: ' + str(ppp_multimatch_sets)])
            # row += 1
            row += 1
            worksheet.write_row(row, 0, ['Identification type M: Manually identified in MS-DIAL'])
            row += 1
            worksheet.write_row(row, 0, ['Identification type T2: Automated identification with RT, MS1, and MS2'])
            row += 1
            worksheet.write_row(row, 0, ['Identification type T1: Automated identification with RT and MS1'])
            row += 1
            worksheet.write_row(row, 0, ['Identification type U: Unknown'])
            row -= 1
            col1 = col = num_peakpair_columns
            start_cell = convert_to_grid(row, col)
            col2 = col = num_total_columns - 1
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            worksheet.merge_range(m_range,  'L/H ratio', merge_format1)
            row += 1
            col1 = col = num_peakpair_columns
            start_cell = convert_to_grid(row, col)
            col2 = col = num_peakpair_columns + num_sample_average_columns + num_sample_signal_columns - 1
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            if (col2 - col1) > 0:
                worksheet.merge_range(m_range,  'samples spiked with heavy-tagged pool', merge_format2)
            else:
                worksheet.write_row(row, col, ['samples spiked with heavy-tagged pool'], merge_format2)
            col += 1
            col1 = col
            start_cell = convert_to_grid(row, col)
            col2 = col = num_total_columns - 1
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            merge_string = 'mix QCs with theoretical L/H ratio ' + str(theoret_mix_lth_ratio)
            if (col2 - col1) > 0:
                worksheet.merge_range(m_range,  merge_string, merge_format3)
            else:
                worksheet.write_row(row, col, [merge_string], merge_format3)
            row += 1
            col1 = col = num_peakpair_columns
            start_cell = convert_to_grid(row, col)
            col2 = col = num_peakpair_columns + num_sample_average_columns - 1
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            if multimatrix:
                if (col2 - col1) > 0:
                    worksheet.merge_range(m_range,  'group average', merge_format4)
                else:
                    worksheet.write_row(row, col, ['group average'], merge_format4)
            col += 1
            col1 = col
            start_cell = convert_to_grid(row, col)
            col2 = col = num_peakpair_columns + num_sample_average_columns + num_sample_signal_columns - 1
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            if multimatrix:
                if (col2 - col1) > 0:
                    worksheet.merge_range(m_range,  'ratio per replicate or sample', merge_format5)
                else:
                    worksheet.write_row(row, col, ['ratio per replicate or sample'], merge_format5)
            else:
                if (col2 - col1) > 0:
                    worksheet.merge_range(m_range,  'ratio per replicate or sample', merge_format2)
                else:
                    worksheet.write_row(row, col, ['ratio per replicate or sample'], merge_format2)
            col += 1
            col1 = col
            start_cell = convert_to_grid(row, col)
            col += num_mix_average_columns - 1
            col2 = col
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            if multimatrix:
                if (col2 - col1) > 0:
                    worksheet.merge_range(m_range,  'group average', merge_format6)
                else:
                    worksheet.write_row(row, col, ['group average'], merge_format6)
            col += 1
            col1 = col
            start_cell = convert_to_grid(row, col)
            col += num_mix_signal_columns - 1
            col2 = col
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            if multimatrix:
                if (col2 - col1) > 0:
                    worksheet.merge_range(m_range,  'ratio per replicate or sample', merge_format7)
                else:
                    worksheet.write_row(row, col, ['ratio per replicate or sample'], merge_format7)
            else:
                if (col2 - col1) > 0:
                    worksheet.merge_range(m_range,  'ratio per replicate or sample', merge_format3)
                else:
                    worksheet.write_row(row, col, ['ratio per replicate or sample'], merge_format3)
            row += 1
            col = 0
            start_cell = convert_to_grid(row, col)
            col = num_peakpair_columns - 1
            end_cell = convert_to_grid(row, col)
            m_range = start_cell + ":" + end_cell
            worksheet.merge_range(m_range,  'metabolite information', merge_format1)
            col += 1
            col += num_sample_average_columns
            for index in final_sample_signal_indexes:
                if multimatrix:
                    worksheet.write_row(row, col, [columns_line[index-1]], merge_format5)
                else:
                    worksheet.write_row(row, col, [columns_line[index-1]], merge_format2)
                col += 1
            col += num_mix_average_columns
            for index in final_mix_signal_indexes:
                if multimatrix:
                    worksheet.write_row(row, col, [columns_line[index-1]], merge_format7)
                else:
                    worksheet.write_row(row, col, [columns_line[index-1]], merge_format3)
                col += 1
            row += 1
            col = 0
            worksheet.write_row(row, col, peakpair_columns, merge_format1)
            col += num_peakpair_columns
            if multimatrix:
                for index in final_sample_average_indexes:
                    worksheet.write_row(row, col, [columns_line[index-1]], merge_format4)
                    col += 1
            for index in final_sample_signal_indexes:
                if multimatrix:
                    worksheet.write_row(row, col, [classes_line[index-1]], merge_format5)
                else:
                    worksheet.write_row(row, col, [classes_line[index-1]], merge_format2)
                col += 1
            if multimatrix:
                for index in final_mix_average_indexes:
                    worksheet.write_row(row, col, [columns_line[index-1]], merge_format6)
                    col += 1
            for index in final_mix_signal_indexes:
                if multimatrix:
                    worksheet.write_row(row, col, [classes_line[index-1]], merge_format7)
                else:
                    worksheet.write_row(row, col, [classes_line[index-1]], merge_format3)
                col += 1
            row += 1

            for peak_pair in final_identified_peak_sets:
                col = 0
                metabolite = peak_pair[0][metabolite_head_pos-1]
                if metabolite == 'Unknown':
                    is_unknown = True
                else:
                    is_unknown = False
                name_frag = metabolite[0:8]
                matched_ms2_string = peak_pair[0][matched_ms2_head_pos-1]
                score_string = peak_pair[0][total_score_head_pos-1]
                if score_string == 'null':
                    manually_assigned = True
                else:
                    manually_assigned = False
                if matched_ms2_string == 'True':
                    matched_ms2 = True
                else:
                    matched_ms2 = False
                wo_ms2 = False
                if name_frag == 'w/o MS2:':
                    wo_ms2 = True
                    metabolite = metabolite[8:]
                if not is_unknown:
                    if manually_assigned:
                        id_type = 'M'
                    elif matched_ms2 and not wo_ms2:
                        id_type = 'T2'
                    else:
                        id_type = 'T1'
                else:
                    id_type = 'U'
                alignment_id = int(peak_pair[0][alignment_id_head_pos-1])
                rt = float(peak_pair[0][rt_head_pos-1])
                light_mz = float(peak_pair[0][mz_head_pos-1])
                heavy_mz = float(peak_pair[1][mz_head_pos-1])
                adduct = peak_pair[0][adduct_head_pos-1]
                tag_count = peak_pair[len(peak_pair)-1]
                if not is_unknown:
                    if peak_pair[0][total_score_head_pos-1] == 'null':
                        score = 'null'
                    else:
                        score = float(peak_pair[0][total_score_head_pos-1])
                else:
                    score = 'null'
                worksheet.write_row(row, col, [metabolite, alignment_id, rt, light_mz, heavy_mz, tag_count, adduct, id_type, score], text_format)
                col += num_peakpair_columns
                if multimatrix:
                    for index in final_sample_average_indexes:
                        light_area = float(peak_pair[0][index-1])
                        heavy_area = float(peak_pair[1][index-1])
                        if heavy_area > 0:
                            lth = light_area / heavy_area
                            worksheet.write_row(row, col, [lth], number_format4)
                        else:
                            lth = 'UND'
                            worksheet.write_row(row, col, [lth], number_format4)
                        col += 1
                for index in final_sample_signal_indexes:
                    light_area = float(peak_pair[0][index-1])
                    heavy_area = float(peak_pair[1][index-1])
                    if heavy_area > 0:
                        lth = light_area / heavy_area
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format5)
                        else:
                            worksheet.write_row(row, col, [lth], number_format2)
                    else:
                        lth = 'UND'
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format5)
                        else:
                            worksheet.write_row(row, col, [lth], number_format2)
                    col += 1
                if multimatrix:
                    for index in final_mix_average_indexes:
                        light_area = float(peak_pair[0][index-1])
                        heavy_area = float(peak_pair[1][index-1])
                        if heavy_area > 0:
                            lth = float(light_area / heavy_area)
                            worksheet.write_row(row, col, [lth], number_format6)
                        else:
                            lth = 'UND'
                            worksheet.write_row(row, col, [lth], number_format6)
                        col += 1
                for index in final_mix_signal_indexes:
                    light_area = float(peak_pair[0][index-1])
                    heavy_area = float(peak_pair[1][index-1])
                    if heavy_area > 0:
                        lth = light_area / heavy_area
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format7)
                        else:
                            worksheet.write_row(row, col, [lth], number_format3)
                    else:
                        lth = 'UND'
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format7)
                        else:
                            worksheet.write_row(row, col, [lth], number_format3)
                    col += 1
                row += 1

            for peak_pair in final_unknown_peak_sets:
                col = 0
                metabolite = peak_pair[0][metabolite_head_pos-1]
                if metabolite == 'Unknown':
                    is_unknown = True
                else:
                    is_unknown = False
                name_frag = metabolite[0:8]
                matched_ms2_string = peak_pair[0][matched_ms2_head_pos-1]
                manual_assignment_string = peak_pair[0][manually_identified_head_pos-1]
                if manual_assignment_string == 'False':
                    manually_assigned = False
                else:
                    manually_assigned = True
                if matched_ms2_string == 'True':
                    matched_ms2 = True
                else:
                    matched_ms2 = False
                wo_ms2 = False
                if name_frag == 'w/o MS2:':
                    wo_ms2 = True
                    metabolite = metabolite[8:]
                if not is_unknown:
                    if manually_assigned:
                        id_type = 'M'
                    elif matched_ms2 and not wo_ms2:
                        id_type = 'T2'
                    else:
                        id_type = 'T1'
                else:
                    id_type = 'U'
                alignment_id = int(peak_pair[0][alignment_id_head_pos-1])
                rt = float(peak_pair[0][rt_head_pos-1])
                light_mz = float(peak_pair[0][mz_head_pos-1])
                heavy_mz = float(peak_pair[1][mz_head_pos-1])
                adduct = peak_pair[0][adduct_head_pos-1]
                tag_count = peak_pair[len(peak_pair) - 1]
                if not is_unknown:
                    score = float(peak_pair[0][total_score_head_pos-1])
                else:
                    score = 'null'
                worksheet.write_row(row, col, [metabolite, alignment_id, rt, light_mz, heavy_mz, tag_count, adduct, id_type, score], text_format)
                col += num_peakpair_columns
                if multimatrix:
                    for index in final_sample_average_indexes:
                        light_area = float(peak_pair[0][index-1])
                        heavy_area = float(peak_pair[1][index-1])
                        if heavy_area > 0:
                            lth = light_area / heavy_area
                            worksheet.write_row(row, col, [lth], number_format4)
                        else:
                            lth = 'UND'
                            worksheet.write_row(row, col, [lth], number_format4)
                        col += 1
                for index in final_sample_signal_indexes:
                    light_area = float(peak_pair[0][index-1])
                    heavy_area = float(peak_pair[1][index-1])
                    if heavy_area > 0:
                        lth = light_area / heavy_area
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format5)
                        else:
                            worksheet.write_row(row, col, [lth], number_format2)
                    else:
                        lth = 'UND'
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format5)
                        else:
                            worksheet.write_row(row, col, [lth], number_format2)
                    col += 1
                if multimatrix:
                    for index in final_mix_average_indexes:
                        light_area = float(peak_pair[0][index-1])
                        heavy_area = float(peak_pair[1][index-1])
                        if heavy_area > 0:
                            lth = float(light_area / heavy_area)
                            worksheet.write_row(row, col, [lth], number_format6)
                        else:
                            lth = 'UND'
                            worksheet.write_row(row, col, [lth], number_format6)
                        col += 1
                for index in final_mix_signal_indexes:
                    light_area = float(peak_pair[0][index-1])
                    heavy_area = float(peak_pair[1][index-1])
                    if heavy_area > 0:
                        lth = light_area / heavy_area
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format7)
                        else:
                            worksheet.write_row(row, col, [lth], number_format3)
                    else:
                        lth = 'UND'
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format7)
                        else:
                            worksheet.write_row(row, col, [lth], number_format3)
                    col += 1
                row += 1

            for peak_pair in multimatch_sets:
                col = 0
                metabolite = peak_pair[0][metabolite_head_pos-1]
                if metabolite == 'Unknown':
                    is_unknown = True
                else:
                    is_unknown = False
                name_frag = metabolite[0:8]
                matched_ms2_string = peak_pair[0][matched_ms2_head_pos-1]
                manual_assignment_string = peak_pair[0][manually_identified_head_pos-1]
                if manual_assignment_string == 'False':
                    manually_assigned = False
                else:
                    manually_assigned = True
                if matched_ms2_string == 'True':
                    matched_ms2 = True
                else:
                    matched_ms2 = False
                wo_ms2 = False
                if name_frag == 'w/o MS2:':
                    wo_ms2 = True
                    metabolite = metabolite[8:]
                if not is_unknown:
                    if manually_assigned:
                        id_type = 'M'
                    elif matched_ms2 and not wo_ms2:
                        id_type = 'T2'
                    else:
                        id_type = 'T1'
                else:
                    id_type = 'U'
                alignment_id = int(peak_pair[0][alignment_id_head_pos-1])
                rt = float(peak_pair[0][rt_head_pos-1])
                light_mz = float(peak_pair[0][mz_head_pos-1])
                heavy_mz = float(peak_pair[1][mz_head_pos-1])
                adduct = peak_pair[0][adduct_head_pos-1]
                if not is_unknown:
                    score = float(peak_pair[0][total_score_head_pos-1])
                else:
                    score = 'null'
                worksheet.write_row(row, col, [metabolite, alignment_id, rt, light_mz, heavy_mz, adduct, id_type, score], text_format)
                col += num_peakpair_columns
                if multimatrix:
                    for index in final_sample_average_indexes:
                        light_area = float(peak_pair[0][index-1])
                        heavy_area = float(peak_pair[1][index-1])
                        if heavy_area > 0:
                            lth = light_area / heavy_area
                            worksheet.write_row(row, col, [lth], number_format4)
                        else:
                            lth = 'UND'
                            worksheet.write_row(row, col, [lth], number_format4)
                        col += 1
                for index in final_sample_signal_indexes:
                    light_area = float(peak_pair[0][index-1])
                    heavy_area = float(peak_pair[1][index-1])
                    if heavy_area > 0:
                        lth = light_area / heavy_area
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format5)
                        else:
                            worksheet.write_row(row, col, [lth], number_format2)
                    else:
                        lth = 'UND'
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format5)
                        else:
                            worksheet.write_row(row, col, [lth], number_format2)
                    col += 1
                if multimatrix:
                    for index in final_mix_average_indexes:
                        light_area = float(peak_pair[0][index-1])
                        heavy_area = float(peak_pair[1][index-1])
                        if heavy_area > 0:
                            lth = float(light_area / heavy_area)
                            worksheet.write_row(row, col, [lth], number_format6)
                        else:
                            lth = 'UND'
                            worksheet.write_row(row, col, [lth], number_format6)
                        col += 1
                for index in final_mix_signal_indexes:
                    light_area = float(peak_pair[0][index-1])
                    heavy_area = float(peak_pair[1][index-1])
                    if heavy_area > 0:
                        lth = light_area / heavy_area
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format7)
                        else:
                            worksheet.write_row(row, col, [lth], number_format3)
                    else:
                        lth = 'UND'
                        if multimatrix:
                            worksheet.write_row(row, col, [lth], number_format7)
                        else:
                            worksheet.write_row(row, col, [lth], number_format3)
                    col += 1
                row += 1
            workbook.close()
        print('')
    return False


def main_function():

    sG.theme('DefaultNoMoreNagging')
    exit_ppp = False
    analyze = False
    area_directory = ''
    mdf_range_bottom = - 500
    mdf_range_top = 499
    light_lth_ratio = 10.0
    heavy_htl_ratio = 100.0
    mix_lth_ratio = 1.0
    mix_lth_tolerance = 0.2
    tag_count_string = '1,2,3'
    tag_counts = [1,2,3]
    tag_isotopic_shift = '0.0000'
    tag_isotopic_shift_tolerance = 10.0
    delete_multimatches = False
    subtract_blank = True
    subtract_natural_heavy = True
    now = str(datetime.now())
    date_str = ''
    for i in now:
        if is_number(i):
            date_str += i
    processed_name = 'name' + date_str
    format_menu_choice = 'Report'
    destination_directory = os.getcwd()

    while not exit_ppp:
        if analyze:
            print('')
            print('Working...')
            print('')
            analyze = call_ppp(area_directory, mdf_range_bottom, mdf_range_top, light_lth_ratio, heavy_htl_ratio, mix_lth_ratio, mix_lth_tolerance, tag_counts, tag_isotopic_shift, tag_isotopic_shift_tolerance, delete_multimatches, subtract_blank, subtract_natural_heavy, processed_name, format_menu_choice, destination_directory, tag_count_string)
            print('')
            print('Finished.')
            print('You may change parameters and GO again.')
            print('')

        layout = [
                    [sG.Text('Alignment Matrix Post Processing for Peak Pair Validation and Ratio Quantification of Isotopic Labeling LC-MS(/MS) Data', font='Arial 11 bold')],
                    [sG.Text('Alignment matrix directory\\name: ' + '\t' + '\t' + '\t' + '\t'), sG.InputText(area_directory), sG.FileBrowse(file_types=(("Text Files", "*.txt"),))],
                    [sG.Text('Mass defect filter lower limit in mDa: ' + '\t' + '\t' + '\t' + '\t'), sG.InputText(mdf_range_bottom)],
                    [sG.Text('Mass defect filter upper limit in mDa: ' + '\t' + '\t' + '\t' + '\t'), sG.InputText(mdf_range_top)],
                    [sG.Text('Minimum light QC L/H ratio: ' + '\t' + '\t' + '\t' + '\t' + '\t'), sG.InputText(light_lth_ratio)],
                    [sG.Text('Minimum heavy QC H/L ratio: ' + '\t' + '\t' + '\t' + '\t'), sG.InputText(heavy_htl_ratio)],
                    [sG.Text('Mix QC theoretical L/H ratio: ' + '\t' + '\t' + '\t' + '\t'), sG.InputText(mix_lth_ratio)],
                    [sG.Text('Mix QC L/H ratio tolerance: ' + '\t' + '\t' + '\t' + '\t' + '\t'), sG.InputText(mix_lth_tolerance)],
                    [sG.Text('Number of tags per molecule (type all for exhaustive search): ' + '\t'), sG.InputText(tag_count_string)],
                    [sG.Text('Exact mass shift between light and heavy tags in Da: ' + '\t' + '\t'), sG.InputText(tag_isotopic_shift)],
                    [sG.Text('Mass shift tolerance in ppm: ' + '\t' + '\t' + '\t' + '\t'), sG.InputText(tag_isotopic_shift_tolerance)],
                    [sG.Text('Subtract background values from samples and QCs?: ' + '\t' + '\t'), sG.Checkbox('Subtract', subtract_blank)],
                    [sG.Text('Subtract isotopic overlap from samples and Mix QCs?: ' + '\t' + '\t'), sG.Checkbox('Subtract', subtract_natural_heavy)],
                    [sG.Text('Output file name: ' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t'), sG.InputText(processed_name)],
                    [sG.Text('Output file format: ' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t'), sG.OptionMenu(['Report', 'Matrix'], format_menu_choice, s=(15, 2)), sG.Text('(Report for summary, matrix for deep dive)')],
                    [sG.Text('Output file directory: ' + '\t' + '\t' + '\t' + '\t' + '\t'), sG.InputText(destination_directory), sG.FolderBrowse()],
                    [sG.Button('GO'), sG.Button('Exit'), sG.Text('This window will disappear while working.')]
                ]

        window = sG.Window('Peak Pair Pruner v1.1', layout)
        while True:
            event, values = window.read()
            if event == sG.WIN_CLOSED or event == 'Exit':
                exit_ppp = True
                break
            if event == 'GO':

                analyze = True

                j = 0
                area_directory = str(values[j])
                if not exists(values[j]):
                    analyze = False
                    print('Error: Alignment matrix not selected or not found.')
                j += 1

                mdf_range_bottom = values[j]
                if not str_is_number_pos_or_neg(values[j]):
                    analyze = False
                    print('Error: mass defect filter lower limit (in mDa) should be a unitless number without spaces.')
                else:
                    mdf_range_bottom = int(round(float(values[j])))
                    if mdf_range_top < - 500:
                        analyze = False
                        print('Error: mass defect filter lower limit (in mDa) should be >= -500.')
                        print('Mass defect is defined as distance from closest nominal mass (integer, Da).')
                        print('Example: -600 mDa from X integer is incorrect. It is really +400 mDa from X-1 integer.')
                j += 1

                mdf_range_top = values[j]
                if not str_is_number_pos_or_neg(values[j]):
                    analyze = False
                    print('Error: mass defect filter upper limit (in mDa) should be a unitless number without spaces.')
                else:
                    mdf_range_top = int(round(float(values[j])))
                    if mdf_range_top > 499:
                        analyze = False
                        print('Error: mass defect filter upper limit (in mDa) should be <= 499.')
                        print('Mass defect is defined as distance from closest nominal mass (integer, Da).')
                        print('For example: +750 mDa from X integer is incorrect. It is really -250 mDa from X+1 integer.')
                j += 1

                if analyze:
                    if mdf_range_top <= mdf_range_bottom:
                        analyze = False
                        print('Error: mass defect filter upper limit must be greater than mass defect filter lower limit.')

                light_lth_ratio = values[j]
                if not str_is_number(values[j]):
                    analyze = False
                    print('Error: minimum light QC L/H ratio should be a unitless number >= 0 without spaces.')
                else:
                    light_lth_ratio = float(values[j])
                j += 1

                heavy_htl_ratio = values[j]
                if not str_is_number(values[j]):
                    analyze = False
                    print('Error: minimum heavy QC H/L ratio should be a unitless number >= 0 without spaces.')
                else:
                    heavy_htl_ratio = float(values[j])
                j += 1

                mix_lth_ratio = values[j]
                if not str_is_number(values[j]):
                    analyze = False
                    print('Error: mix QC theoretical L/H ratio should be a unitless number >= 0 without spaces.')
                else:
                    mix_lth_ratio = float(values[j])
                j += 1

                if str_is_number((values[j])):
                    mix_lth_tolerance = float(values[j])
                if not str_is_number(values[j]):
                    analyze = False
                    print('Error: mix QC ratio tolerance should be a unitless number >= 0 without spaces.')
                elif analyze and mix_lth_ratio <= mix_lth_tolerance:
                    analyze = False
                    print('Error: mix QC ratio tolerance should be less mix QC theoretical L/H ratio.')
                else:
                    mix_lth_tolerance = float(values[j])
                j += 1

                tag_count_string = values[j]
                tag_counts = []
                current_tag = ''
                add_tag = False
                if tag_count_string == 'all' or tag_count_string == 'All' or tag_count_string == 'ALL':
                    tag_counts = ['all']
                else:
                    if tag_count_string[0] == ',':
                        analyze = False
                        print('Error: Number of tags per molecule must have an integer before the first comma.')
                    elif analyze:
                        for char in tag_count_string:
                            if char == ',':
                                add_tag = True
                            elif not str_is_integer(char):
                                analyze = False
                                print('Error: Number of tags per molecule must be one or more positive integers separated by commas.')
                            else:
                                current_tag += char
                            if add_tag and analyze:
                                if current_tag[0] == '0':
                                    print('Error: Number of tags per molecule should be a positive integer without leading zeroes.')
                                    analyze = False
                                else:
                                    tag_counts += [int(current_tag)]
                                    add_tag = False
                                    current_tag = ''
                        if analyze:
                            tag_counts += [int(current_tag)]
                j += 1

                tag_isotopic_shift = values[j]
                if not str_is_number(values[j]):
                    analyze = False
                    print('Error: exact mass shift between light and heavy tags (in Da) should be a unitless number without spaces.')
                elif float(values[j]) < 1:
                    analyze = False
                    print('Error: exact mass shift between light and heavy tags (in Da) should be >= 1.')
                else:
                    tag_isotopic_shift = float(values[j])
                    if tag_isotopic_shift == 0:
                        analyze = False
                        print('Error: exact mass shift between light and heavy tags (in Da) should be a unitless number without spaces.')
                j += 1

                tag_isotopic_shift_tolerance = values[j]
                if not str_is_number(values[j]):
                    analyze = False
                    print('Error: mass shift tolerance should be a unitless number >= 0 without spaces.')
                else:
                    tag_isotopic_shift_tolerance = float(values[j])
                j += 1

                # A multi-match is a set of 3 or more peaks with overlapping pairings (a theoretical alignment error).
                # Multi-matches have not been encountered and so have not been thoroughly tested.
                # As processing multi-matches has not been thoroughly tested, they will be deleted if encountered.
                delete_multimatches = True
                subtract_blank = bool(values[j])
                j += 1
                subtract_natural_heavy = bool(values[j])
                j += 1

                processed_name = str(values[j])
                directory_and_name = destination_directory + '/' + processed_name + '.xlsx'
                if os.path.isfile(directory_and_name):
                    analyze = False
                    now = str(datetime.now())
                    date_str = ''
                    for i in now:
                        if is_number(i):
                            date_str += i
                    processed_name = processed_name + date_str
                    print('Error: a file already exists with your output file name.')
                    print('Random number string appended to output file name.')
                    print('Press GO again or first change the output file name to something more desirable.')
                j += 1

                format_menu_choice = str(values[j])
                if format_menu_choice == 'Matrix' and len(tag_counts) > 1:
                    analyze = False
                    print('Error: Matrix output only supports one tag count per job.')
                j += 1
                destination_directory = str(values[j])
                j += 1

                window.close()
                break
    window.close()


main_function()
