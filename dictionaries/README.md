# Dictionaries

This section contains hands-on exercises for understanding and working with Python dictionaries.

## Exercise: Dictionary Product Lookup

**File:** `dictionary_product_lookup.py`

This exercise uses a dictionary to represent a small shop inventory. Each product is stored as a **key**, and its price is stored as the corresponding **value**.

### What the program does

1. Stores products and their prices in a dictionary.
2. Asks the user to enter a product name.
3. Checks whether the product exists in the dictionary.
4. If the product exists, displays and speaks its price.
5. If the product does not exist, asks whether the user wants to add it.
6. If the user chooses `yes`, asks for the price and adds the new product to the dictionary.

### Python concepts practiced

- Creating dictionaries
- Key-value pairs
- Accessing dictionary values using keys
- Checking whether a key exists using `in`
- Adding a new key-value pair
- `input()` for user interaction
- `if`, `else` and nested conditions
- f-strings
- Type conversion with `int()`

### How the dictionary works

The main data structure follows this pattern:

```python
shop = {
    "Laptop": 55000,
    "Monitor": 12000,
    "Mouse": 800
}
```

Here:

- `"Laptop"`, `"Monitor"`, and `"Mouse"` are **keys**.
- `55000`, `12000`, and `800` are their corresponding **values**.

To check whether a product exists:

```python
if stock in shop:
```

To retrieve its price:

```python
shop[stock]
```

To add a new product:

```python
shop[stock] = new_price
```

## Text-to-Speech

The exercise uses **Windows SAPI** through `win32com.client` for text-to-speech:

```python
import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")
```

The program then uses `speaker.Speak()` to read messages aloud.

### Alternative

`pyttsx3` is another Python text-to-speech library that can be used for a similar feature. It can be a useful alternative when learning different approaches to text-to-speech.

The current exercise uses `win32com.client` because it works with the Windows SAPI speech system and is the implementation used in this exercise.

## Requirements

- Python 3
- Windows (for the current `win32com.client` implementation)
- `pywin32` package

Install the required package with:

```bash
pip install pywin32
```

## Example

```text
Enter your product : Laptop
The price of Laptop is 55000
```

If the product is not found:

```text
Enter your product : Camera
Do you want to add this product? (yes/no): yes
Enter the price: 5000
```

The new product is then added to the `shop` dictionary for the current program run.

> **Note:** The dictionary is stored in memory, so products added while the program is running are not permanently saved after the program exits.
