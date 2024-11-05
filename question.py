import pyautogui
import time

def ask_question():
    # Display the question in a UI using PyAutoGUI
    pyautogui.alert("Is HJ a slave master?", "Quiz")

    # Wait for the user to answer
    time.sleep(1)

    # Check the user's answer
    answer = pyautogui.confirm("Select your answer:", "Quiz", buttons=["Yes", "No"])

    if answer == "Yes":
        pyautogui.alert("Congratulations, Right Answer", "Quiz")
    elif answer == "No":
        pyautogui.alert("Wrong answer, select again", "Quiz")
        ask_question()
    else:
        pyautogui.alert("Invalid choice. Please select either 'Yes' or 'No'.", "Quiz")
        ask_question()

if __name__ == "__main__":
    ask_question()
