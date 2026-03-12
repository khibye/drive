from docxtpl import DocxTemplate

doc = DocxTemplate("template.docx")

hebrew_letters = "אבגדהוזחטיכלמנסעפצקרשת"

actions = [
    {"title": "לסיים דוח", "owner": "hana", "deadline": "today"},
    {"title": "לשלוח מייל", "owner": "dana", "deadline": "tomorrow"},
]

context = {
    "actions": [
        {
            "number":   f"{i+1}. {action['title']}",
            "owner":    f"{hebrew_letters[0]}) {action['owner']}",
            "deadline": f"{hebrew_letters[1]}) {action['deadline']}",
        }
        for i, action in enumerate(actions)
    ]
}

doc.render(context)
doc.save("output.docx")


{%p for action in actions %}
{{ action.number }}
{{ action.owner }}
{{ action.deadline }}
{%p endfor %}
