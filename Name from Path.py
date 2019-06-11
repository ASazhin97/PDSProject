def name_from_path(path):
    name_comma_index = None
    for i in range(len(path)):
        if path[i] == ',':
            name_comma_index = i
    for j in range(name_comma_index, 1, -1):
        if path[j] == "\\":
            lastname = path[j+1:name_comma_index]
            break
    for h in range(name_comma_index, len(path)):
        if path[h] == "\\":
            firstname = path[name_comma_index + 2:h]
            break
    fullname = firstname + " " + lastname
    return fullname