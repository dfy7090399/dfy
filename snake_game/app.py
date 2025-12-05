from flask import Flask, render_template, jsonify, request
import random
import json

app = Flask(__name__)

class SnakeGame:
    def __init__(self, width=20, height=20):
        self.width = width
        self.height = height
        self.reset_game()
    
    def reset_game(self):
        # Snake starts in the middle
        self.snake = [(self.width//2, self.height//2)]
        self.direction = 'right'
        self.food = self.generate_food()
        self.game_over = False
        self.score = 0
    
    def generate_food(self):
        while True:
            food = (random.randint(0, self.width-1), random.randint(0, self.height-1))
            if food not in self.snake:
                return food
    
    def move_snake(self, direction):
        if self.game_over:
            return False
        
        # Prevent snake from going back into itself
        opposite_directions = {
            'up': 'down', 'down': 'up', 'left': 'right', 'right': 'left'
        }
        if direction == opposite_directions.get(self.direction):
            return False
        
        self.direction = direction
        head_x, head_y = self.snake[0]
        
        # Calculate new head position
        if self.direction == 'up':
            new_head = (head_x, head_y - 1)
        elif self.direction == 'down':
            new_head = (head_x, head_y + 1)
        elif self.direction == 'left':
            new_head = (head_x - 1, head_y)
        elif self.direction == 'right':
            new_head = (head_x + 1, head_y)
        
        # Check wall collision
        if (new_head[0] < 0 or new_head[0] >= self.width or 
            new_head[1] < 0 or new_head[1] >= self.height):
            self.game_over = True
            return False
        
        # Check self collision
        if new_head in self.snake:
            self.game_over = True
            return False
        
        # Move snake
        self.snake.insert(0, new_head)
        
        # Check if food eaten
        if new_head == self.food:
            self.score += 10
            self.food = self.generate_food()
        else:
            self.snake.pop()
        
        return True
    
    def get_state(self):
        return {
            'snake': self.snake,
            'food': self.food,
            'game_over': self.game_over,
            'score': self.score,
            'direction': self.direction
        }

# Global game instance
game = SnakeGame()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/game_state')
def get_game_state():
    return jsonify(game.get_state())

@app.route('/api/move', methods=['POST'])
def move_snake():
    data = request.get_json()
    direction = data.get('direction')
    success = game.move_snake(direction)
    return jsonify({'success': success, 'state': game.get_state()})

@app.route('/api/reset', methods=['POST'])
def reset_game():
    game.reset_game()
    return jsonify(game.get_state())

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
