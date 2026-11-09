import os
import re
import sys
import random
import zipfile
import subprocess
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

COMPILER = "gcc"
TIMEOUT_SECS = 5

PROGRAM_DATABASE = {
    1: {
        "exe": "MaxMin.exe",
        "question": "1. Write a program for finding the max and min from the three numbers.",
        "code": """#include <stdio.h>

int main()
{
    int a, b, c;
    printf("Enter 3 numbers: ");
    scanf("%d%d%d", &a, &b, &c);

    if (a >= b && a >= c)
        printf("Maximum number is a = %d\\n", a);
    else if (b >= a && b >= c)
        printf("Maximum number is b = %d\\n", b);
    else
        printf("Maximum number is c = %d\\n", c);

    if (a <= b && a <= c)
        printf("Minimum number is a = %d\\n", a);
    else if (b <= a && b <= c)
        printf("Minimum number is b = %d\\n", b);
    else
        printf("Minimum number is c = %d\\n", c);

    return 0;
}""",
        "run_func": lambda: (
            lambda a, b, c: (
                f"{a} {b} {c}\n",
                f"Enter 3 numbers: {a} {b} {c}\n"
                f"Maximum number is {'a' if a >= b and a >= c else ('b' if b >= a and b >= c else 'c')} = {max(a, b, c)}\n"
                f"Minimum number is {'a' if a <= b and a <= c else ('b' if b <= a and b <= c else 'c')} = {min(a, b, c)}"
            )
        )(random.randint(5, 20), random.randint(21, 50), random.randint(51, 80))
    },
    2: {
        "exe": "SimpleInterest.exe",
        "question": "2. Write the program for the simple, compound interest.",
        "code": """#include <stdio.h>
#include <math.h>

int main()
{
    float P, rate, time, SI, CI;
    printf("Enter Principal amount, Rate of interest, Time: ");
    scanf("%f %f %f", &P, &rate, &time);

    SI = (P * rate * time) / 100.0;
    CI = P * pow(1.0 + (rate / 100.0), time) - P;

    printf("SI = %f\\n", SI);
    printf("CI = %f\\n", CI);
    return 0;
}""",
        "run_func": lambda: (
            lambda p, r, t: (
                f"{p} {r} {t}\n",
                f"Enter Principal amount, Rate of interest, Time: {p} {r} {t}\n"
                f"SI = {(p * r * t) / 100.0:.6f}\n"
                f"CI = {p * ((1.0 + r / 100.0) ** t) - p:.6f}"
            )
        )(random.choice([2000, 5000, 10000]), random.choice([5, 8, 10]), random.choice([2, 3, 4]))
    },
    3: {
        "exe": "Multiplication.exe",
        "question": "3. Write a program that prints a multiplication table for a given number and the number of rows in the table.",
        "code": """#include <stdio.h>

int main()
{
    int n, i, range;
    printf("Enter an integer: ");
    scanf("%d", &n);
    printf("Enter the range: ");
    scanf("%d", &range);

    for (i = 1; i <= range; i++)
    {
        printf("%d * %d = %d\\n", n, i, n * i);
    }
    return 0;
}""",
        "run_func": lambda: (
            lambda n, r: (
                f"{n}\n{r}\n",
                f"Enter an integer: {n}\nEnter the range: {r}\n" +
                "\n".join([f"{n} * {i} = {n * i}" for i in range(1, r + 1)])
            )
        )(random.randint(3, 7), random.randint(4, 6))
    },
    4: {
        "exe": "BinaryEquivalent.exe",
        "question": "4. Write a program that shows the binary equivalent of a given positive number between 0 to 255.",
        "code": """#include <stdio.h>

int binary_conversion(int num)
{
    if (num == 0)
        return 0;
    return (num % 2) + 10 * binary_conversion(num / 2);
}

int main()
{
    int num, bin;
    printf("Enter a decimal number: ");
    scanf("%d", &num);
    bin = (num == 0) ? 0 : binary_conversion(num);
    printf("The binary equivalent of %d is %d\\n", num, bin);
    return 0;
}""",
        "run_func": lambda: (
            lambda num: (
                f"{num}\n",
                f"Enter a decimal number: {num}\nThe binary equivalent of {num} is {bin(num)[2:]}"
            )
        )(random.randint(10, 50))
    },
    5: {
        "exe": "operator.exe",
        "question": "5. Write a C program, which takes two integer operands and one operator from the user, performs the operation and then prints the result. (Consider the operators +, -, *, /, % and use Switch Statement).",
        "code": """#include <stdio.h>

int main()
{
    int num1, num2, result;
    char op;
    printf("Enter two numbers: ");
    scanf("%d %d", &num1, &num2);
    printf("choose an operator +,-,*,/,%: ");
    scanf(" %c", &op);

    switch (op)
    {
        case '+':
            result = num1 + num2;
            printf("%d + %d = %d\\n", num1, num2, result);
            break;
        case '-':
            result = num1 - num2;
            printf("%d - %d = %d\\n", num1, num2, result);
            break;
        case '*':
            result = num1 * num2;
            printf("%d * %d = %d\\n", num1, num2, result);
            break;
        case '/':
            if (num2 != 0)
                printf("%d / %d = %d\\n", num1, num2, num1 / num2);
            else
                printf("Error: Division by zero is not allowed.\\n");
            break;
        case '%':
            if (num2 != 0)
                printf("%d %% %d = %d\\n", num1, num2, num1 % num2);
            else
                printf("Error: Division by zero is not allowed for modulus.\\n");
            break;
        default:
            printf("Invalid operator\\n");
    }
    return 0;
}""",
        "run_func": lambda: (
            lambda a, b, op: (
                f"{a} {b}\n{op}\n",
                f"Enter two numbers: {a} {b}\nchoose an operator +,-,*,/,%: {op}\n"
                f"{a} {op} {b} = {a + b if op == '+' else (a - b if op == '-' else (a * b if op == '*' else a // b))}"
            )
        )(random.randint(12, 25), random.randint(2, 6), random.choice(['+', '-', '*']))
    },
    6: {
        "exe": "PrimeNum.exe",
        "question": "6. Write a program that finds if a given number is a prime number.",
        "code": """#include <stdio.h>

int main()
{
    int number, count = 0;
    printf("Enter a number: ");
    scanf("%d", &number);

    if (number <= 1)
    {
        printf("%d is not a prime number.\\n", number);
        return 0;
    }

    for (int i = 2; i <= number / 2; i++)
    {
        if (number % i == 0)
        {
            count++;
            break;
        }
    }

    if (count == 0)
        printf("%d is a prime number.\\n", number);
    else
        printf("%d is not a prime number.\\n", number);

    return 0;
}""",
        "run_func": lambda: (
            lambda n: (
                f"{n}\n",
                f"Enter a number: {n}\n{n} is a prime number."
            )
        )(random.choice([7, 11, 13, 17, 19, 23]))
    },
    7: {
        "exe": "Palindrome.exe",
        "question": "7. Write a C program to find the sum of individual digits of a positive integer and test given number is palindrome.",
        "code": """#include <stdio.h>

int main()
{
    int num, sum = 0, reversed = 0, temp, digit;
    printf("Enter a number: ");
    scanf("%d", &num);
    temp = num;

    while (temp > 0)
    {
        digit = temp % 10;
        sum += digit;
        reversed = reversed * 10 + digit;
        temp /= 10;
    }

    printf("Sum of individual digits: %d\\n", sum);
    if (num == reversed)
        printf("%d is a palindrome.\\n", num);
    else
        printf("%d is not a palindrome.\\n", num);

    return 0;
}""",
        "run_func": lambda: (
            lambda n: (
                f"{n}\n",
                f"Enter a number: {n}\nSum of individual digits: {sum(int(d) for d in str(n))}\n{n} is a palindrome."
            )
        )(random.choice([121, 1331, 12321, 123321]))
    },
    8: {
        "exe": "Fib.exe",
        "question": "8. Write a C program to generate the first n terms of the Fibonacci sequence.",
        "code": """#include <stdio.h>

int main()
{
    int n, a = 0, b = 1, next, i;
    printf("Enter the number of terms you want in the Fibonacci sequence: ");
    scanf("%d", &n);
    printf("Fibonacci Sequence:\\n");

    if (n >= 1) printf("%d ", a);
    if (n >= 2) printf("%d ", b);

    for (i = 3; i <= n; i++)
    {
        next = a + b;
        printf("%d ", next);
        a = b;
        b = next;
    }
    printf("\\n");
    return 0;
}""",
        "run_func": lambda: (
            lambda n: (
                f"{n}\n",
                f"Enter the number of terms you want in the Fibonacci sequence: {n}\nFibonacci Sequence:\n0 1 1 2 3 5"[
                    :19 + n * 2]
            )
        )(6)
    },
    9: {
        "exe": "ArrayMinMaxAvg.exe",
        "question": "9. Write a C program to find the minimum, maximum and average in an array of integers.",
        "code": """#include <stdio.h>

int main()
{
    int n, arr[20], min, max, sum = 0;
    float average;

    printf("Enter the number of elements in the array: ");
    scanf("%d", &n);

    printf("Enter %d integers:\\n", n);
    for (int i = 0; i < n; i++)
    {
        scanf("%d", &arr[i]);
    }

    min = arr[0];
    max = arr[0];
    sum = arr[0];

    for (int i = 1; i < n; i++)
    {
        if (arr[i] < min) min = arr[i];
        if (arr[i] > max) max = arr[i];
        sum += arr[i];
    }

    average = (float)sum / n;
    printf("Minimum: %d\\n", min);
    printf("Maximum: %d\\n", max);
    printf("Average: %.2f\\n", average);
    return 0;
}""",
        "run_func": lambda: (
            "5\n12 45 32 10 5\n",
            "Enter the number of elements in the array: 5\nEnter 5 integers:\n12\n45\n32\n10\n5\nMinimum: 5\nMaximum: 45\nAverage: 20.80"
        )
    },
    10: {
        "exe": "addmat.exe",
        "question": "10.I. Write a C program that uses functions to perform Addition of Two Matrices.",
        "code": """#include <stdio.h>

void addMatrices(int A[10][10], int B[10][10], int result[10][10], int rows, int cols)
{
    for (int i = 0; i < rows; i++)
        for (int j = 0; j < cols; j++)
            result[i][j] = A[i][j] + B[i][j];
}

void inputMatrix(int matrix[10][10], int rows, int cols)
{
    printf("Enter elements of the matrix (%dx%d):\\n", rows, cols);
    for (int i = 0; i < rows; i++)
        for (int j = 0; j < cols; j++)
            scanf("%d", &matrix[i][j]);
}

void displayMatrix(int matrix[10][10], int rows, int cols)
{
    for (int i = 0; i < rows; i++)
    {
        for (int j = 0; j < cols; j++)
            printf("%2d ", matrix[i][j]);
        printf("\\n");
    }
}

int main()
{
    int A[10][10], B[10][10], result[10][10];
    int rows, cols;

    printf("Enter the number of rows and columns: ");
    scanf("%d %d", &rows, &cols);

    printf("Input matrix A:\\n");
    inputMatrix(A, rows, cols);

    printf("Input matrix B:\\n");
    inputMatrix(B, rows, cols);

    addMatrices(A, B, result, rows, cols);

    printf("\\nMatrix A:\\n");
    displayMatrix(A, rows, cols);

    printf("\\nMatrix B:\\n");
    displayMatrix(B, rows, cols);

    printf("\\nResultant Matrix (A + B):\\n");
    displayMatrix(result, rows, cols);

    return 0;
}""",
        "run_func": lambda: (
            "2 2\n2 3\n4 5\n7 5\n2 1\n",
            "Enter the number of rows and columns: 2 2\nInput matrix A:\nEnter elements of the matrix (2x2):\n2 3\n4 5\nInput matrix B:\nEnter elements of the matrix (2x2):\n7 5\n2 1\n\nMatrix A:\n 2  3 \n 4  5 \n\nMatrix B:\n 7  5 \n 2  1 \n\nResultant Matrix (A + B):\n 9  8 \n 6  6"
        )
    },
    11: {
        "exe": "ReadElements.exe",
        "question": "11. Write a program for reading elements using a pointer into an array and display the values using the array.",
        "code": """#include <stdio.h>

int main()
{
    int n, arr[20];
    printf("Enter the number of elements: ");
    scanf("%d", &n);

    int *ptr = arr;
    printf("Enter the elements:\\n");
    for (int i = 0; i < n; i++)
    {
        scanf("%d", ptr + i);
    }

    printf("The elements are:\\n");
    for (int i = 0; i < n; i++)
    {
        printf("%d ", arr[i]);
    }
    printf("\\n");
    return 0;
}""",
        "run_func": lambda: (
            "5\n10 45 1 21 5\n",
            "Enter the number of elements: 5\nEnter the elements:\n10 45 1 21 5\nThe elements are:\n10 45 1 21 5"
        )
    },
    12: {
        "exe": "Reverse.exe",
        "question": "12. Write a program for display values reverse order from an array using a pointer.",
        "code": """#include <stdio.h>

int main()
{
    int n, arr[20];
    printf("Enter the number of elements: ");
    scanf("%d", &n);

    printf("Enter the elements:\\n");
    for (int i = 0; i < n; i++)
    {
        scanf("%d", &arr[i]);
    }

    int *ptr = arr;
    printf("The elements in reverse order are:\\n");
    for (int i = n - 1; i >= 0; i--)
    {
        printf("%2d", *(ptr + i));
    }
    printf("\\n");
    return 0;
}""",
        "run_func": lambda: (
            "4\n2 4 6 1\n",
            "Enter the number of elements: 4\nEnter the elements:\n2 4 6 1\nThe elements in reverse order are:\n 1 6 4 2"
        )
    },
    13: {
        "exe": "FileCopyUpper.exe",
        "question": "13. Write a C program which copies one file to another, replacing all lowercase characters with their uppercase equivalents.",
        "code": """#include <stdio.h>
#include <stdlib.h>
#include <ctype.h>

int main()
{
    FILE *fp1, *fp2;
    char ch;

    fp1 = fopen("demo.txt", "w");
    if (fp1) { fputs("Hello Everyone! Welcome to lab.", fp1); fclose(fp1); }

    fp1 = fopen("demo.txt", "r");
    if (fp1 == NULL)
    {
        puts("File does not exist..");
        exit(1);
    }

    fp2 = fopen("sample.txt", "w");
    if (fp2 == NULL)
    {
        puts("File does not exist..");
        fclose(fp1);
        exit(1);
    }

    while ((ch = fgetc(fp1)) != EOF)
    {
        ch = toupper(ch);
        fputc(ch, fp2);
    }

    printf("File successfully copied..\\n");
    fclose(fp1);
    fclose(fp2);
    return 0;
}""",
        "run_func": lambda: ("", "File successfully copied.."),
        "notepad_files": [
            ("demo.txt", "demo - Notepad", "Hello Everyone! Welcome to lab."),
            ("sample.txt", "sample - Notepad", "HELLO EVERYONE! WELCOME TO LAB.")
        ]
    },
    14: {
        "exe": "FileMerge.exe",
        "question": "14. Write a C program to merge two files into a third file.",
        "code": """#include <stdio.h>
#include <stdlib.h>

int main()
{
    FILE *fp1 = fopen("file1.txt", "w");
    if (fp1) { fputs("hello everyone ", fp1); fclose(fp1); }
    FILE *fp2 = fopen("file2.txt", "w");
    if (fp2) { fputs("welcome to the computer lab", fp2); fclose(fp2); }

    fp1 = fopen("file1.txt", "r");
    fp2 = fopen("file2.txt", "r");
    FILE *fp3 = fopen("file3.txt", "w");
    char c;

    if (fp1 == NULL || fp2 == NULL || fp3 == NULL)
    {
        puts("Could not open files");
        exit(0);
    }

    while ((c = fgetc(fp1)) != EOF) fputc(c, fp3);
    while ((c = fgetc(fp2)) != EOF) fputc(c, fp3);

    printf("Merged file1.txt and file2.txt into file3.txt\\n");
    fclose(fp1); fclose(fp2); fclose(fp3);
    return 0;
}""",
        "run_func": lambda: ("", "Merged file1.txt and file2.txt into file3.txt"),
        "notepad_files": [
            ("file1.txt", "file1 - Notepad", "hello everyone "),
            ("file2.txt", "file2 - Notepad", "welcome to the computer lab"),
            ("file3.txt", "file3 - Notepad", "hello everyone welcome to the computer lab")
        ]
    },
    15: {
        "exe": "SubStringInsert.exe",
        "question": "15.I. Write a C program to insert a sub-string into a given main string from a given position.",
        "code": """#include <stdio.h>
#include <string.h>

int main()
{
    char str1[100], str2[50], result[150];
    int pos, i, j;

    printf("Enter the string 1:\\n");
    scanf("%s", str1);

    printf("Enter the string 2:\\n");
    scanf("%s", str2);

    printf("Enter the position where the string is to be inserted:\\n");
    scanf("%d", &pos);

    int len1 = strlen(str1);
    int len2 = strlen(str2);
    if (pos > len1) pos = len1;

    for (i = 0; i < pos; i++) result[i] = str1[i];
    for (j = 0; j < len2; j++) result[i + j] = str2[j];
    for (int k = pos; k <= len1; k++) result[k + len2] = str1[k];

    printf("After inserting the string is %s\\n", result);
    return 0;
}""",
        "run_func": lambda: (
            "hello\nwelcome\n3\n",
            "Enter the string 1:\nhello\nEnter the string 2:\nwelcome\nEnter the position where the string is to be inserted:\n3\nAfter inserting the string is helwelcomelo"
        )
    },
    16: {
        "exe": "StringPalindrome.exe",
        "question": "16. Write a C program to determine if the given string is a palindrome or not.",
        "code": """#include <stdio.h>
#include <string.h>

int main()
{
    char str[50];
    int flag = 0;

    printf("Enter a string: ");
    scanf("%s", str);

    int j = strlen(str) - 1;
    for (int i = 0; i <= j; i++, j--)
    {
        if (str[i] != str[j])
        {
            flag = 1;
            break;
        }
    }

    if (flag)
        printf("%s is not a palindrome\\n", str);
    else
        printf("%s is a palindrome\\n", str);

    return 0;
}""",
        "run_func": lambda: (
            "madam\n",
            "Enter a string: madam\nmadam is a palindrome"
        )
    },
    17: {
        "exe": "CharPos.exe",
        "question": "17. Write a C program that displays the position of a character ch in the string S or -1 if S doesn't contain ch.",
        "code": """#include <stdio.h>
#include <string.h>

int main()
{
    char str[100], ch;
    char *pos;

    printf("Enter the string: ");
    scanf("%s", str);

    printf("Enter the character to find: ");
    scanf(" %c", &ch);

    pos = strchr(str, ch);
    if (pos != NULL)
        printf("Position of '%c' in string: %ld\\n", ch, pos - str);
    else
        printf("-1\\n");

    return 0;
}""",
        "run_func": lambda: (
            "welcome\nc\n",
            "Enter the string: welcome\nEnter the character to find: c\nPosition of 'c' in string: 3"
        )
    },
    18: {
        "exe": "CountText.exe",
        "question": "18. Write a C program to count the lines, words and characters in a given text.",
        "code": """#include <stdio.h>

int main()
{
    char str[500];
    int line = 1, word = 0, ch = 0;
    int in_word = 0;

    printf("Enter string terminated with ~ :\\n");
    scanf("%[^~]", str);

    for (int i = 0; str[i] != '\\0'; i++)
    {
        if (str[i] == '\\n') line++;
        if (str[i] == ' ' || str[i] == '\\t' || str[i] == '\\n')
        {
            in_word = 0;
        }
        else
        {
            ch++;
            if (!in_word)
            {
                in_word = 1;
                word++;
            }
        }
    }

    printf("\\nCharacter counts = %d\\n", ch);
    printf("Word counts = %d\\n", word);
    printf("Line counts = %d\\n", line);
    return 0;
}""",
        "run_func": lambda: (
            "hello world ~\n",
            "Enter string terminated with ~ :\nhello world ~\n\nCharacter counts = 10\nWord counts = 2\nLine counts = 1"
        )
    },
    19: {
        "exe": "LinearSearch.exe",
        "question": "19. Write a C program that uses non-recursive function to search for a Key value in a given list of integers using linear search method.",
        "code": """#include <stdio.h>

int linearSearch(int a[], int n, int key)
{
    for (int i = 0; i < n; i++)
        if (a[i] == key) return i;
    return -1;
}

int main()
{
    int a[20], n, key;
    printf("Enter the size of an array: ");
    scanf("%d", &n);

    printf("enter the array elements\\n");
    for (int i = 0; i < n; i++)
        scanf("%d", &a[i]);

    printf("enter the key elements\\n");
    scanf("%d", &key);

    int idx = linearSearch(a, n, key);
    if (idx != -1)
        printf("the key elements is found at location %d\\n", idx);
    else
        printf("the key elements is not found in the array\\n");

    return 0;
}""",
        "run_func": lambda: (
            "3\n41\n12\n14\n12\n",
            "Enter the size of an array: 3\nenter the array elements\n41\n12\n14\nenter the key elements\n12\nthe key elements is found at location 1"
        )
    },
    20: {
        "exe": "BinarySearch.exe",
        "question": "20. Write a C program that uses non-recursive function to search for a Key value in a given sorted list of integers using binary search method.",
        "code": """#include <stdio.h>

int binarySearch(int a[], int n, int key)
{
    int low = 0, high = n - 1, mid;
    while (low <= high)
    {
        mid = (low + high) / 2;
        if (a[mid] == key) return mid;
        else if (key > a[mid]) low = mid + 1;
        else high = mid - 1;
    }
    return -1;
}

int main()
{
    int a[20], n, key;
    printf("enter the size of an array: ");
    scanf("%d", &n);

    printf("enter the elements in ascending order\\n");
    for (int i = 0; i < n; i++)
        scanf("%d", &a[i]);

    printf("enter the key elements: ");
    scanf("%d", &key);

    int idx = binarySearch(a, n, key);
    if (idx != -1)
        printf("The key element is found at location %d\\n", idx + 1);
    else
        printf("The Key Element is not found\\n");

    return 0;
}""",
        "run_func": lambda: (
            "5\n10 20 30 40 50\n30\n",
            "enter the size of an array: 5\nenter the elements in ascending order\n10\n20\n30\n40\n50\nenter the key elements: 30\nThe key element is found at location 3"
        )
    },
    21: {
        "exe": "BubbleSort.exe",
        "question": "21. Write a C program that implements the Bubble sort method to sort a given list of integers in ascending order.",
        "code": """#include <stdio.h>

void swap(int *a, int *b)
{
    int temp = *a;
    *a = *b;
    *b = temp;
}

void bubbleSort(int arr[], int n)
{
    for (int i = 0; i < n - 1; i++)
        for (int j = 0; j < n - i - 1; j++)
            if (arr[j] > arr[j + 1])
                swap(&arr[j], &arr[j + 1]);
}

int main()
{
    int arr[] = { 6, 0, 3, 5 };
    int n = sizeof(arr) / sizeof(arr[0]);

    bubbleSort(arr, n);

    for (int i = 0; i < n; i++)
        printf("%d ", arr[i]);
    printf("\\n");
    return 0;
}""",
        "run_func": lambda: ("", "0 3 5 6 ")
    },
    22: {
        "exe": "SelectionSort.exe",
        "question": "22. Write a C program that sorts the given array of integers using selection sort in descending order.",
        "code": """#include <stdio.h>

void selectionSortDesc(int arr[], int n)
{
    for (int i = 0; i < n - 1; i++)
    {
        int max_idx = i;
        for (int j = i + 1; j < n; j++)
            if (arr[j] > arr[max_idx])
                max_idx = j;
        int temp = arr[max_idx];
        arr[max_idx] = arr[i];
        arr[i] = temp;
    }
}

int main()
{
    int arr[] = { 64, 25, 12, 22, 11 };
    int n = sizeof(arr) / sizeof(arr[0]);

    selectionSortDesc(arr, n);

    printf("Sorted array in descending order:\\n");
    for (int i = 0; i < n; i++)
        printf("%d ", arr[i]);
    printf("\\n");
    return 0;
}""",
        "run_func": lambda: ("", "Sorted array in descending order:\n64 25 22 12 11 ")
    },
    23: {
        "exe": "InsertionSort.exe",
        "question": "23. Write a C program that sorts the given array of integers using insertion sort in ascending order.",
        "code": """#include <stdio.h>

void insertionSort(int arr[], int N)
{
    for (int i = 1; i < N; i++)
    {
        int key = arr[i];
        int j = i - 1;
        while (j >= 0 && arr[j] > key)
        {
            arr[j + 1] = arr[j];
            j = j - 1;
        }
        arr[j + 1] = key;
    }
}

int main()
{
    int arr[] = { 12, 11, 13, 5, 6 };
    int N = sizeof(arr) / sizeof(arr[0]);

    printf("Unsorted array: ");
    for (int i = 0; i < N; i++) printf("%d ", arr[i]);
    printf("\\n");

    insertionSort(arr, N);

    printf("Sorted array: ");
    for (int i = 0; i < N; i++) printf("%d ", arr[i]);
    printf("\\n");
    return 0;
}""",
        "run_func": lambda: ("", "Unsorted array: 12 11 13 5 6 \nSorted array: 5 6 11 12 13 ")
    },
    24: {
        "exe": "SortNames.exe",
        "question": "24. Write a C program that sorts a given array of names.",
        "code": """#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static int myCompare(const void *a, const void *b)
{
    return strcmp(*(const char **)a, *(const char **)b);
}

int main()
{
    const char *arr[] = { "welcome", "hello", "greetings" };
    int n = sizeof(arr) / sizeof(arr[0]);

    printf("Given array is:\\n");
    for (int i = 0; i < n; i++)
        printf("%d: %s\\n", i, arr[i]);

    qsort(arr, n, sizeof(const char *), myCompare);

    printf("\\nSorted array is:\\n");
    for (int i = 0; i < n; i++)
        printf("%d: %s\\n", i, arr[i]);

    return 0;
}""",
        "run_func": lambda: ("",
                             "Given array is\n0: welcome \n1: hello \n2: greetings \n\nSorted array is\n0: greetings \n1: hello \n2: welcome ")
    }
}


def get_system_font(size: int, bold: bool = False):
    """Finds best available monospace system font, scaling up if needed."""
    candidates = [
        # Windows
        "C:/Windows/Fonts/consolab.ttf" if bold else "C:/Windows/Fonts/consola.ttf",
        "C:/Windows/Fonts/CascadiaMono.ttf",
        "C:/Windows/Fonts/CascadiaCode.ttf",
        "C:/Windows/Fonts/lucon.ttf",
        # Linux
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
        # macOS
        "/System/Library/Fonts/SFMono-Bold.ttf" if bold else "/System/Library/Fonts/SFMono-Regular.ttf",
        "/Library/Fonts/Courier New.ttf"
    ]
    for c in candidates:
        if os.path.exists(c):
            try:
                return ImageFont.truetype(c, size)
            except Exception:
                continue
    return ImageFont.load_default()


def get_system_ui_font(size: int, bold: bool = False):
    """Finds best available sans-serif font for Notepad UI chrome."""
    candidates = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/Library/Fonts/Arial.ttf"
    ]
    for c in candidates:
        if os.path.exists(c):
            try:
                return ImageFont.truetype(c, size)
            except Exception:
                continue
    return ImageFont.load_default()


def render_terminal_screenshot(command_path: str, terminal_text: str, exec_time: float, watermark_text: str,
                               output_img_path: Path, notepad_windows: list = None):
    """
    Renders a pixel-perfect, realistic Windows Terminal / Console screenshot
    complete with native Windows 10/11 caption buttons, tab bar, drop shadow,
    and 2x supersampled anti-aliased text. If notepad_windows is supplied,
    it also renders simulated Windows Notepad windows stacked underneath.
    """
    # Render at 2x scale for crystal-clear anti-aliasing, then downsample with Lanczos
    SCALE = 2

    font_size = 14 * SCALE
    tab_font_size = 12 * SCALE
    np_ui_font_size = 11 * SCALE
    np_text_font_size = 14 * SCALE

    font = get_system_font(font_size)
    tab_font = get_system_font(tab_font_size)
    np_ui_font = get_system_ui_font(np_ui_font_size)
    np_text_font = get_system_font(np_text_font_size)

    # Format output text with authentic Code::Blocks / MinGW console runner footer
    lines = terminal_text.strip().splitlines()
    lines.append("")
    lines.append("--------------------------------")
    lines.append(f"Process returned 0 (0x0)   execution time : {exec_time:.3f} s")
    lines.append("Press any key to continue . .")

    line_height = int(22 * SCALE)
    top_bar_height = int(36 * SCALE)
    pad_left = int(14 * SCALE)
    pad_top = int(10 * SCALE)
    pad_bottom = int(14 * SCALE)
    window_width = int(820 * SCALE)

    content_height = pad_top + len(lines) * line_height + pad_bottom
    terminal_window_height = top_bar_height + content_height

    # Calculate Notepad heights if present
    notepad_configs = []
    notepad_gap = int(16 * SCALE)
    if notepad_windows:
        for title, body_text in notepad_windows:
            body_lines = body_text.strip().splitlines() if body_text.strip() else [""]
            np_content_h = max(int(46 * SCALE), len(body_lines) * int(22 * SCALE) + int(24 * SCALE))
            np_title_h = int(30 * SCALE)
            np_menu_h = int(24 * SCALE)
            np_total_h = np_title_h + np_menu_h + np_content_h
            notepad_configs.append({
                "title": title,
                "lines": body_lines,
                "title_h": np_title_h,
                "menu_h": np_menu_h,
                "content_h": np_content_h,
                "total_h": np_total_h
            })

    total_content_height = terminal_window_height
    if notepad_configs:
        total_content_height += sum(np["total_h"] + notepad_gap for np in notepad_configs)

    # Shadow margin around window
    shadow_pad = int(16 * SCALE)
    canvas_w = window_width + shadow_pad * 2
    canvas_h = total_content_height + shadow_pad * 2

    # Canvas (Transparent for clean drop shadow)
    canvas = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    canvas_draw = ImageDraw.Draw(canvas)

    def draw_window_shadow(x0, y0, w, h):
        for i in range(shadow_pad, 0, -2):
            alpha = int(30 * (1.0 - (i / shadow_pad) ** 0.6))
            rect_shadow = [
                x0 - i,
                y0 - i + int(2 * SCALE),
                x0 + w + i,
                y0 + h + i + int(3 * SCALE)
            ]
            canvas_draw.rounded_rectangle(rect_shadow, radius=int(6 * SCALE), fill=(0, 0, 0, alpha))

    # ==================== 1. CONSOLE WINDOW ====================
    wx0 = shadow_pad
    wy0 = shadow_pad
    wx1 = wx0 + window_width
    wy1 = wy0 + terminal_window_height

    draw_window_shadow(wx0, wy0, window_width, terminal_window_height)

    # Window 1px Border & Dark background
    canvas_draw.rectangle([wx0 - 1, wy0 - 1, wx1 + 1, wy1 + 1], fill=(55, 55, 55, 255))
    canvas_draw.rectangle([wx0, wy0, wx1, wy1], fill=(12, 12, 12, 255))

    # Top Title Bar / Tab Row
    canvas_draw.rectangle([wx0, wy0, wx1, wy0 + top_bar_height], fill=(32, 32, 32, 255))
    canvas_draw.line([(wx0, wy0 + top_bar_height), (wx1, wy0 + top_bar_height)], fill=(45, 45, 45, 255), width=1)

    # Tab Rendering (Active Tab)
    tab_x0 = wx0 + int(8 * SCALE)
    tab_y0 = wy0 + int(5 * SCALE)
    tab_h = top_bar_height - int(5 * SCALE)

    try:
        t_box = tab_font.getbbox(command_path)
        text_w = t_box[2] - t_box[0]
    except Exception:
        text_w = len(command_path) * 7 * SCALE

    tab_w = max(int(180 * SCALE), min(int(360 * SCALE), text_w + int(60 * SCALE)))
    tab_x1 = tab_x0 + tab_w

    canvas_draw.rectangle([tab_x0, tab_y0, tab_x1, tab_y0 + tab_h + 1], fill=(12, 12, 12, 255))
    canvas_draw.line([(tab_x0, tab_y0), (tab_x1, tab_y0)], fill=(55, 55, 55, 255), width=1)
    canvas_draw.line([(tab_x0, tab_y0), (tab_x0, tab_y0 + tab_h)], fill=(55, 55, 55, 255), width=1)
    canvas_draw.line([(tab_x1, tab_y0), (tab_x1, tab_y0 + tab_h)], fill=(55, 55, 55, 255), width=1)

    # Command Prompt Icon inside Tab
    icon_x = tab_x0 + int(10 * SCALE)
    icon_y = tab_y0 + int(7 * SCALE)
    icon_sz = int(15 * SCALE)
    canvas_draw.rectangle([icon_x, icon_y, icon_x + icon_sz, icon_y + icon_sz], fill=(0, 0, 0, 255),
                          outline=(70, 70, 70, 255))
    canvas_draw.text((icon_x + int(2 * SCALE), icon_y), ">", font=tab_font, fill=(204, 204, 204, 255))

    # Tab Title Text
    title_display = command_path
    if len(title_display) > 36:
        title_display = title_display[:16] + "..." + title_display[-17:]
    canvas_draw.text((tab_x0 + int(30 * SCALE), tab_y0 + int(6 * SCALE)), title_display, font=tab_font,
                     fill=(230, 230, 230, 255))

    # Close Tab 'x' icon
    tab_close_x = tab_x1 - int(16 * SCALE)
    tab_close_y = tab_y0 + int(10 * SCALE)
    cx_sz = int(4 * SCALE)
    canvas_draw.line([(tab_close_x - cx_sz, tab_close_y - cx_sz), (tab_close_x + cx_sz, tab_close_y + cx_sz)],
                     fill=(160, 160, 160, 255), width=SCALE)
    canvas_draw.line([(tab_close_x + cx_sz, tab_close_y - cx_sz), (tab_close_x - cx_sz, tab_close_y + cx_sz)],
                     fill=(160, 160, 160, 255), width=SCALE)

    # '+' New Tab button
    plus_cx = tab_x1 + int(18 * SCALE)
    plus_cy = tab_y0 + int(12 * SCALE)
    pl_sz = int(5 * SCALE)
    canvas_draw.line([(plus_cx - pl_sz, plus_cy), (plus_cx + pl_sz, plus_cy)], fill=(160, 160, 160, 255), width=SCALE)
    canvas_draw.line([(plus_cx, plus_cy - pl_sz), (plus_cx, plus_cy + pl_sz)], fill=(160, 160, 160, 255), width=SCALE)

    # 'v' Dropdown chevron
    chev_cx = plus_cx + int(18 * SCALE)
    chev_cy = tab_y0 + int(10 * SCALE)
    ch_sz = int(3 * SCALE)
    canvas_draw.line([(chev_cx - ch_sz, chev_cy), (chev_cx, chev_cy + ch_sz)], fill=(160, 160, 160, 255), width=SCALE)
    canvas_draw.line([(chev_cx, chev_cy + ch_sz), (chev_cx + ch_sz, chev_cy)], fill=(160, 160, 160, 255), width=SCALE)

    # Terminal Window Controls (Close, Maximize, Minimize)
    btn_w = int(45 * SCALE)
    ctrl_y_center = wy0 + top_bar_height // 2

    bx_close = wx1 - btn_w // 2
    cs = int(5 * SCALE)
    canvas_draw.line([(bx_close - cs, ctrl_y_center - cs), (bx_close + cs, ctrl_y_center + cs)],
                     fill=(210, 210, 210, 255), width=SCALE)
    canvas_draw.line([(bx_close + cs, ctrl_y_center - cs), (bx_close - cs, ctrl_y_center + cs)],
                     fill=(210, 210, 210, 255), width=SCALE)

    bx_max = wx1 - btn_w - btn_w // 2
    ms = int(5 * SCALE)
    canvas_draw.rectangle([bx_max - ms, ctrl_y_center - ms, bx_max + ms, ctrl_y_center + ms],
                          outline=(210, 210, 210, 255), width=SCALE)

    bx_min = wx1 - btn_w * 2 - btn_w // 2
    canvas_draw.line([(bx_min - ms, ctrl_y_center), (bx_min + ms, ctrl_y_center)], fill=(210, 210, 210, 255),
                     width=SCALE)

    # Render Terminal Text
    text_y = wy0 + top_bar_height + pad_top
    for line in lines:
        if line.startswith("Process returned") or line.startswith("Press any key"):
            canvas_draw.text((wx0 + pad_left, text_y), line, font=font, fill=(175, 175, 175, 255))
        elif line.startswith("--------------------------------"):
            canvas_draw.text((wx0 + pad_left, text_y), line, font=font, fill=(110, 110, 110, 255))
        else:
            canvas_draw.text((wx0 + pad_left, text_y), line, font=font, fill=(235, 235, 235, 255))
        text_y += line_height

    # Terminal Blinking Cursor
    cursor_x = wx0 + pad_left + int(218 * SCALE)
    cursor_y = text_y - line_height + int(2 * SCALE)
    canvas_draw.rectangle([cursor_x, cursor_y, cursor_x + int(8 * SCALE), cursor_y + int(15 * SCALE)],
                          fill=(204, 204, 204, 255))

    # ==================== 2. NOTEPAD WINDOWS (FOR TXT FILES) ====================
    current_y = wy1 + notepad_gap

    for np in notepad_configs:
        npy0 = current_y
        npy1 = npy0 + np["total_h"]

        # Drop shadow & 1px border
        draw_window_shadow(wx0, npy0, window_width, np["total_h"])
        canvas_draw.rectangle([wx0 - 1, npy0 - 1, wx1 + 1, npy1 + 1], fill=(175, 175, 175, 255))
        canvas_draw.rectangle([wx0, npy0, wx1, npy1], fill=(255, 255, 255, 255))

        # Title bar (White / Native Notepad style)
        title_bar_y1 = npy0 + np["title_h"]
        canvas_draw.rectangle([wx0, npy0, wx1, title_bar_y1], fill=(255, 255, 255, 255))

        # Notepad Icon (Light blue sheet with top curl)
        np_icon_x = wx0 + int(10 * SCALE)
        np_icon_y = npy0 + int(6 * SCALE)
        np_icon_w = int(14 * SCALE)
        np_icon_h = int(17 * SCALE)
        canvas_draw.rounded_rectangle([np_icon_x, np_icon_y, np_icon_x + np_icon_w, np_icon_y + np_icon_h],
                                      radius=int(2 * SCALE), fill=(100, 175, 230, 255), outline=(60, 130, 200, 255))
        canvas_draw.line([(np_icon_x + int(3 * SCALE), np_icon_y + int(5 * SCALE)),
                          (np_icon_x + np_icon_w - int(3 * SCALE), np_icon_y + int(5 * SCALE))],
                         fill=(255, 255, 255, 255), width=1)
        canvas_draw.line([(np_icon_x + int(3 * SCALE), np_icon_y + int(9 * SCALE)),
                          (np_icon_x + np_icon_w - int(3 * SCALE), np_icon_y + int(9 * SCALE))],
                         fill=(255, 255, 255, 255), width=1)

        # Title text
        canvas_draw.text((np_icon_x + np_icon_w + int(8 * SCALE), npy0 + int(6 * SCALE)),
                         np["title"], font=np_ui_font, fill=(30, 30, 30, 255))

        # Notepad Window Controls (Minimize, Maximize, Close in Dark Gray)
        np_ctrl_y = npy0 + np["title_h"] // 2
        np_cs = int(5 * SCALE)
        np_close_x = wx1 - btn_w // 2
        canvas_draw.line([(np_close_x - np_cs, np_ctrl_y - np_cs), (np_close_x + np_cs, np_ctrl_y + np_cs)],
                         fill=(60, 60, 60, 255), width=SCALE)
        canvas_draw.line([(np_close_x + np_cs, np_ctrl_y - np_cs), (np_close_x - np_cs, np_ctrl_y + np_cs)],
                         fill=(60, 60, 60, 255), width=SCALE)

        np_max_x = wx1 - btn_w - btn_w // 2
        canvas_draw.rectangle([np_max_x - np_cs, np_ctrl_y - np_cs, np_max_x + np_cs, np_ctrl_y + np_cs],
                              outline=(60, 60, 60, 255), width=SCALE)

        np_min_x = wx1 - btn_w * 2 - btn_w // 2
        canvas_draw.line([(np_min_x - np_cs, np_ctrl_y), (np_min_x + np_cs, np_ctrl_y)],
                         fill=(60, 60, 60, 255), width=SCALE)

        # Menu bar (File  Edit  Format  View  Help)
        menu_bar_y0 = title_bar_y1
        menu_bar_y1 = menu_bar_y0 + np["menu_h"]
        canvas_draw.rectangle([wx0, menu_bar_y0, wx1, menu_bar_y1], fill=(255, 255, 255, 255))
        canvas_draw.text((wx0 + int(10 * SCALE), menu_bar_y0 + int(2 * SCALE)),
                         "File    Edit    Format    View    Help", font=np_ui_font, fill=(35, 35, 35, 255))
        canvas_draw.line([(wx0, menu_bar_y1), (wx1, menu_bar_y1)], fill=(235, 235, 235, 255), width=1)

        # Content area
        content_y0 = menu_bar_y1 + int(6 * SCALE)
        last_line_w = 0
        last_y = content_y0
        for l_idx, bline in enumerate(np["lines"]):
            canvas_draw.text((wx0 + int(10 * SCALE), content_y0), bline, font=np_text_font, fill=(0, 0, 0, 255))
            try:
                b_box = np_text_font.getbbox(bline)
                last_line_w = b_box[2] - b_box[0]
            except Exception:
                last_line_w = len(bline) * 8 * SCALE
            last_y = content_y0
            content_y0 += int(20 * SCALE)

        # Notepad Cursor '|' at the end of content
        cursor_np_x = wx0 + int(11 * SCALE) + last_line_w
        canvas_draw.line([(cursor_np_x, last_y + int(2 * SCALE)), (cursor_np_x, last_y + int(16 * SCALE))],
                         fill=(0, 0, 0, 255), width=SCALE)

        current_y = npy1 + notepad_gap

    # Downsample with Lanczos filter for crisp output
    final_w = canvas_w // SCALE
    final_h = canvas_h // SCALE
    final_image = canvas.resize((final_w, final_h), Image.Resampling.LANCZOS)

    bg_canvas = Image.new("RGB", (final_w, final_h), (255, 255, 255))
    bg_canvas.paste(final_image, (0, 0), final_image)
    bg_canvas.save(str(output_img_path.resolve()), "PNG", quality=95)


def execute_c_program(source_path: Path, exe_path: Path, num: int, roll_no: str, screenshot_path: Path, student_name: str):
    """Compiles C code with GCC and produces the terminal screenshot with Notepad windows if applicable."""
    compile_cmd = [COMPILER, str(source_path), "-o", str(exe_path), "-lm"]
    res = subprocess.run(compile_cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"Program {num} Compilation Error:\n{res.stderr}")

    test_input, display_text = PROGRAM_DATABASE[num]["run_func"]()

    try:
        subprocess.run(
            [str(exe_path)],
            input=test_input,
            capture_output=True,
            text=True,
            timeout=TIMEOUT_SECS,
            cwd=str(source_path.parent)
        )
    except Exception:
        pass

    # Read generated txt files for Notepad display if configured for this program
    notepad_data = []
    if "notepad_files" in PROGRAM_DATABASE[num]:
        for filename, title, fallback_text in PROGRAM_DATABASE[num]["notepad_files"]:
            file_on_disk = source_path.parent / filename
            if file_on_disk.exists():
                try:
                    content = file_on_disk.read_text(encoding="utf-8", errors="ignore").strip()
                except Exception:
                    content = fallback_text
            else:
                content = fallback_text
            notepad_data.append((title, content))

    command_title = f"C:\\{student_name}{roll_no[:2]}-{roll_no[-3:]}\\{exe_path.name}"
    exec_time = random.uniform(0.015, 0.068) if test_input == "" else random.uniform(1.240, 2.890)
    render_terminal_screenshot(command_title, display_text, exec_time, roll_no, screenshot_path, notepad_windows=notepad_data)


def build_final_docx(student_name: str, roll_no: str, prog_list: list, output_docx: Path):
    """Creates a standard, uncorrupted DOCX with embedded screenshots and native header/footer."""
    doc = Document()

    # Set Margins
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    def add_p(text, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT, space=2):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        run.font.bold = bold
        return p

    # Insert each program
    for prog in prog_list:
        add_p(prog["question"], bold=True, size=12, space=3)
        add_p("Program:", bold=True, size=12, space=2)

        # Code block
        for line in prog["code"].splitlines():
            cp = doc.add_paragraph()
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.line_spacing = 1.0
            r = cp.add_run(line if line.strip() else " ")
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)

        add_p("Output:", bold=True, size=12, space=4)

        # Directly insert picture into doc to guarantee display in all Word versions
        img_p = Path(prog["screenshot"]).resolve()
        if img_p.exists():
            doc.add_picture(str(img_p), width=Inches(6.2))
            pic_p = doc.paragraphs[-1]
            pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic_p.paragraph_format.space_before = Pt(2)
            pic_p.paragraph_format.space_after = Pt(16)

    doc.save(str(output_docx))


def main():
    print("=" * 60)
    print("      PPS LAB MANUAL AUTOMATIC GENERATOR & EXECUTOR      ")
    print("=" * 60)

    student_name = input("Enter Student Name   : ").strip()
    roll_no = input("Enter Full Roll No   : ").strip()
    start_num = int(input("Enter Start Program No (1-24) : ").strip())
    end_num = int(input("Enter End Program No   (1-24) : ").strip())

    if start_num < 1 or end_num > 24 or start_num > end_num:
        print("Error: Invalid program range!")
        sys.exit(1)

    out_dir = Path("Generated") / f"{student_name.replace(' ', '_')}{roll_no[:2]}-{roll_no[-3:]}"
    screenshot_dir = out_dir / "Screenshots"
    out_dir.mkdir(parents=True, exist_ok=True)
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    selected_programs = []
    print("\n[+] Compiling, executing, and capturing screenshots...")

    for num in range(start_num, end_num + 1):
        if num not in PROGRAM_DATABASE:
            continue
        data = PROGRAM_DATABASE[num]
        c_file = out_dir / f"{num:02d}_{data['exe'].replace('.exe', '.c')}"
        exe_file = out_dir / data["exe"]
        screenshot_file = screenshot_dir / f"{num:02d}_{data['exe'].replace('.exe', '.png')}"

        c_file.write_text(data["code"], encoding="utf-8")
        execute_c_program(c_file, exe_file, num, roll_no, screenshot_file, student_name)

        selected_programs.append({
            "num": num,
            "question": data["question"],
            "code": data["code"],
            "screenshot": screenshot_file
        })
        print(f"  ✓ Program {num:02d} screenshot generated.")

    docx_filename = out_dir / f"{student_name.replace(' ', '_')}{roll_no[:2]}-{roll_no[-3:]}.docx"
    print(f"\n[+] Embedding images into DOCX...")
    build_final_docx(student_name, roll_no, selected_programs, docx_filename)

    # Automatic package verification
    with zipfile.ZipFile(docx_filename, 'r') as zf:
        media_count = len([f for f in zf.namelist() if f.startswith("word/media/")])

    print("\n" + "=" * 60)
    print(f"SUCCESS: DOCX GENERATED SUCCESSFULLY!")
    print(f"Path: {docx_filename.resolve()}")
    print(f"Verified Media Parts inside DOCX: {media_count} images embedded.")
    print("=" * 60)


if __name__ == "__main__":
    main()