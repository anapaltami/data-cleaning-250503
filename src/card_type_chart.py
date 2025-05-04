import pandas as pd
import matplotlib.pyplot as plt

# load cleaned data
df = pd.read_csv('output/cleaned_data.csv')

# count each card type
card_counts = df['CARD TYPE'].value_counts().reset_index()
card_counts.columns = ['Card Type', 'Count']

# define Catppuccin Mocha colors
background_color = '#1e1e2e'
text_color = '#cdd6f4'
bar_colors = {
    'VISA': '#cba6f7',       # mauve
    'MASTERCARD': '#89b4fa', # blue
    'DISCOVER': '#f2cdcd',   # flamingo
    'AMEX': '#fab387'        # peach
}

# match bar colors to card types
colors = [bar_colors.get(ct, '#ffffff') for ct in card_counts['Card Type']]

# plot bar chart
plt.figure(figsize=(6, 4), facecolor=background_color)
bars = plt.bar(card_counts['Card Type'], card_counts['Count'], color=colors)

# set text and axes styling
plt.title('Card Type Distribution', color=text_color, fontsize=14)
plt.xlabel('Card Type', color=text_color)
plt.ylabel('Number of Users', color=text_color)
plt.xticks(color=text_color)
plt.yticks(color=text_color)
plt.gca().set_facecolor(background_color)
plt.gca().spines['bottom'].set_color(text_color)
plt.gca().spines['left'].set_color(text_color)
plt.gca().spines['top'].set_color(background_color)
plt.gca().spines['right'].set_color(background_color)
plt.tight_layout()

# save the chart
plt.savefig('output/card_type_chart.png', dpi=300, facecolor=background_color)
plt.close()
print("Card type chart saved to output/card_type_chart.png")