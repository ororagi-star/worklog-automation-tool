from copy import deepcopy

from automate_worklog import normalize_text
from worklog_config import clean_item, validate_config


def add_category(config, name):
    draft = deepcopy(config)
    draft["categories"].append(normalize_text(name))
    return validate_config(draft)


def rename_category(config, old, new):
    draft = deepcopy(config)
    new = normalize_text(new)
    draft["categories"] = [new if name == old else name for name in draft["categories"]]
    for course in draft["courses"]:
        if course["category"] == old:
            course["category"] = new
    for template in draft["templates"]:
        for item in template["items"]:
            if item["category"] == old:
                item["category"] = new
    return validate_config(draft)


def delete_category(config, name):
    if any(course["category"] == name for course in config["courses"]):
        raise ValueError("이 카테고리에 수업이 있습니다. 수업의 카테고리를 먼저 변경하거나 수업을 삭제해 주세요.")
    draft = deepcopy(config)
    draft["categories"].remove(name)
    return validate_config(draft)


def save_course(config, values):
    draft = deepcopy(config)
    course = clean_item(values)
    course["course_id"] = course["id"]
    existing = next((index for index, row in enumerate(draft["courses"]) if row["id"] == course["id"]), None)
    if existing is None:
        draft["courses"].append(course)
    else:
        draft["courses"][existing] = course
    for template in draft["templates"]:
        for index, item in enumerate(template["items"]):
            if item["course_id"] == course["id"]:
                template["items"][index] = {**course, "id": item["id"], "course_id": course["id"], "order": item["order"]}
    return validate_config(draft)


def delete_course(config, course_id):
    if any(item["course_id"] == course_id for template in config["templates"] for item in template["items"]):
        raise ValueError("업무일지에서 사용 중입니다. 업무일지 설정에서 먼저 선택을 해제해 주세요.")
    draft = deepcopy(config)
    draft["courses"] = [course for course in draft["courses"] if course["id"] != course_id]
    return validate_config(draft)


def assign_courses(config, template_key, course_ids):
    draft = deepcopy(config)
    template = next(template for template in draft["templates"] if template["key"] == template_key)
    old_items = {item["course_id"]: item for item in template["items"]}
    catalogue = {course["id"]: course for course in draft["courses"]}
    if len(course_ids) != len(set(course_ids)) or any(key not in catalogue for key in course_ids):
        raise ValueError("업무일지에 넣을 수업 목록을 확인해 주세요.")
    template["items"] = [
        {**catalogue[key], "id": old_items[key]["id"] if key in old_items else f"{template_key}:{key}",
         "course_id": key, "order": index}
        for index, key in enumerate(course_ids, 1)
    ]
    return validate_config(draft)


def place_course(config, values, template_keys):
    draft = deepcopy(config)
    category = normalize_text(values.get("category"))
    if category not in draft["categories"]:
        draft = add_category(draft, category)
    draft = save_course(draft, values)
    course_id = values.get("id") or draft["courses"][-1]["id"]
    if not set(template_keys) <= {row["key"] for row in draft["templates"]}:
        raise ValueError("업무일지를 확인해 주세요.")
    for template in draft["templates"]:
        items = template["items"]
        existing = next((item for item in items if item["course_id"] == course_id), None)
        if template["key"] not in template_keys:
            template["items"] = [item for item in items if item["course_id"] != course_id]
        elif existing is None:
            course = next(row for row in draft["courses"] if row["id"] == course_id)
            items.append({**course, "id": f"{template['key']}:{course_id}", "course_id": course_id,
                          "order": max((item["order"] for item in items), default=0) + 1})
    return validate_config(draft)


def remove_course_everywhere(config, course_id):
    draft = deepcopy(config)
    for template in draft["templates"]:
        template["items"] = [item for item in template["items"] if item["course_id"] != course_id]
    draft["courses"] = [course for course in draft["courses"] if course["id"] != course_id]
    return validate_config(draft)
