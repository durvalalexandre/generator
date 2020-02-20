import io


def create_copy_from_default(default_config_file, copy_name):
    # Create a copy from a default parameters file
    with io.open(default_config_file, 'r', encoding='utf-8') as f:
        with io.open(copy_name, "w", encoding='utf-8') as new_file:
            for line in f:
                new_file.write(line)
            new_file.close()


def get_dates_pattern(days, hours):
    days_pattern = []
    if hours is not None:
        all_days = ['seg', 'ter', 'qua', 'qui', 'sex', 'sab', 'dom']
        hours = hours.replace(' ', '')
        if days == 'seg a sex':
            all_days.remove('sab')
            all_days.remove('dom')
        elif days == 'seg a sab':
            all_days.remove('dom')
        elif days == '24h':
            hours = "00:00-24:00"
        else:
            hours = ""

        if 'seg' in all_days:
            days_pattern.append({'pattern1': 'W//1/{}\n'.format(hours)})
        if 'ter' in all_days:
            days_pattern.append({'pattern2': 'W//2/{}\n'.format(hours)})
        if 'qua' in all_days:
            days_pattern.append({'pattern3': 'W//3/{}\n'.format(hours)})
        if 'qui' in all_days:
            days_pattern.append({'pattern4': 'W//4/{}\n'.format(hours)})
        if 'sex' in all_days:
            days_pattern.append({'pattern5': 'W//5/{}\n'.format(hours)})
        if 'sab' in all_days:
            days_pattern.append({'pattern6': 'W//6/{}\n'.format(hours)})
        if 'dom' in all_days:
            days_pattern.append({'pattern7': 'W//0/{}\n'.format(hours)})
    return days_pattern


def normalize_value(option_value):
    if type(option_value) is str:
        if option_value.lower() in ['não', 'nao']:
            return 'false'
        elif option_value.lower() == 'sim':
            return 'true'
        return option_value.strip()
    return int(option_value)
