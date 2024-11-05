import time
import pywinauto

# Function to scroll chat container to load more messages
def scroll_chat_container(chat_container):
    chat_container.scroll(direction="bottom")
    # Wait for a brief moment for messages to load
    time.sleep(1)

# Function to extract chat text from message controls
def extract_chat_text(messages):
    chat_text = []
    for message in messages:
        text = message.window_text()
        if text:  # Exclude empty messages
            chat_text.append(text)
    return chat_text

# Main function to extract chat text from Microsoft Teams application
def extract_teams_chat_text():
    teams_title = "Chat | Sheng Yang Ng | Microsoft Teams"
    app = pywinauto.Application().connect(title=teams_title)
    main_window = app.window(title=teams_title)

    # Assuming the chat messages are displayed in a list or a container
    chat_container = main_window.child_window(class_name="ChatContainerClass")

    # Initial extraction of chat text
    scroll_chat_container(chat_container)
    messages = chat_container.children(class_name="MessageClass")
    chat_text = extract_chat_text(messages)

    # Scroll and extract additional chat text until no more messages can be loaded
    while True:
        try:
            scroll_chat_container(chat_container)
            messages = chat_container.children(class_name="MessageClass")
            new_chat_text = extract_chat_text(messages)
            if new_chat_text:
                chat_text.extend(new_chat_text)
            else:
                break  # No more messages loaded
        except pywinauto.timings.TimeoutError:
            break  # Timeout waiting for messages to load

    return chat_text

# Extract chat text from Microsoft Teams and print it
chat_text = extract_teams_chat_text()
print("Extracted Chat Text:")
for text in chat_text:
    print(text)
