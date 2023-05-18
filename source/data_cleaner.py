
def clean_bullet_points(points):
    data = []
    for line in points.split("\n"):
        if line.startswith(("-", "●", "・")):
            data.append(line.replace("-", "", 1).replace("●", "", 1).replace("・", "", 1).strip())
        else:
            
            
            data.append(line.strip())
    return data