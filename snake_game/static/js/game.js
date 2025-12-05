class SnakeGameClient {
    constructor() {
        this.canvas = document.getElementById('game-canvas');
        this.ctx = this.canvas.getContext('2d');
        this.scoreElement = document.getElementById('score');
        this.finalScoreElement = document.getElementById('final-score');
        this.gameOverElement = document.getElementById('game-over');
        this.resetButton = document.getElementById('reset-btn');
        this.restartButton = document.getElementById('restart-btn');
        
        this.gridSize = 20;
        this.cellSize = this.canvas.width / this.gridSize;
        
        this.setupEventListeners();
        this.gameLoop();
    }
    
    setupEventListeners() {
        // Keyboard controls
        document.addEventListener('keydown', (e) => {
            if (this.gameState && this.gameState.game_over) {
                return;
            }
            
            const keyMap = {
                'ArrowUp': 'up',
                'ArrowDown': 'down', 
                'ArrowLeft': 'left',
                'ArrowRight': 'right',
                'w': 'up',
                's': 'down',
                'a': 'left',
                'd': 'right'
            };
            
            const direction = keyMap[e.key];
            if (direction) {
                e.preventDefault();
                this.moveSnake(direction);
            }
        });
        
        // Button controls
        this.resetButton.addEventListener('click', () => this.resetGame());
        this.restartButton.addEventListener('click', () => {
            this.gameOverElement.classList.add('hidden');
            this.resetGame();
        });
    }
    
    async moveSnake(direction) {
        try {
            const response = await fetch('/api/move', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ direction: direction })
            });
            
            const data = await response.json();
            if (data.success) {
                this.gameState = data.state;
                this.updateDisplay();
            }
        } catch (error) {
            console.error('Error moving snake:', error);
        }
    }
    
    async resetGame() {
        try {
            const response = await fetch('/api/reset', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            });
            
            this.gameState = await response.json();
            this.updateDisplay();
        } catch (error) {
            console.error('Error resetting game:', error);
        }
    }
    
    async fetchGameState() {
        try {
            const response = await fetch('/api/game_state');
            this.gameState = await response.json();
            this.updateDisplay();
        } catch (error) {
            console.error('Error fetching game state:', error);
        }
    }
    
    updateDisplay() {
        if (!this.gameState) return;
        
        // Update score
        this.scoreElement.textContent = this.gameState.score;
        
        // Check game over
        if (this.gameState.game_over) {
            this.finalScoreElement.textContent = this.gameState.score;
            this.gameOverElement.classList.remove('hidden');
        }
        
        // Clear canvas
        this.ctx.fillStyle = 'rgba(0, 0, 0, 0.1)';
        this.ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);
        
        // Draw grid (subtle)
        this.ctx.strokeStyle = 'rgba(255, 255, 255, 0.05)';
        this.ctx.lineWidth = 1;
        for (let i = 0; i <= this.gridSize; i++) {
            const pos = i * this.cellSize;
            this.ctx.beginPath();
            this.ctx.moveTo(pos, 0);
            this.ctx.lineTo(pos, this.canvas.height);
            this.ctx.stroke();
            
            this.ctx.beginPath();
            this.ctx.moveTo(0, pos);
            this.ctx.lineTo(this.canvas.width, pos);
            this.ctx.stroke();
        }
        
        // Draw snake
        this.gameState.snake.forEach((segment, index) => {
            const gradient = this.ctx.createRadialGradient(
                segment[0] * this.cellSize + this.cellSize / 2,
                segment[1] * this.cellSize + this.cellSize / 2,
                0,
                segment[0] * this.cellSize + this.cellSize / 2,
                segment[1] * this.cellSize + this.cellSize / 2,
                this.cellSize / 2
            );
            
            if (index === 0) {
                // Head - brighter green
                gradient.addColorStop(0, '#4ade80');
                gradient.addColorStop(1, '#22c55e');
            } else {
                // Body - gradient fade
                gradient.addColorStop(0, '#22c55e');
                gradient.addColorStop(1, '#16a34a');
            }
            
            this.ctx.fillStyle = gradient;
            this.ctx.fillRect(
                segment[0] * this.cellSize + 2,
                segment[1] * this.cellSize + 2,
                this.cellSize - 4,
                this.cellSize - 4
            );
            
            // Add shine effect
            this.ctx.fillStyle = 'rgba(255, 255, 255, 0.2)';
            this.ctx.fillRect(
                segment[0] * this.cellSize + 4,
                segment[1] * this.cellSize + 4,
                this.cellSize / 3,
                this.cellSize / 3
            );
        });
        
        // Draw food (pulsing effect)
        const time = Date.now() / 1000;
        const pulse = Math.sin(time * 3) * 0.2 + 0.8;
        
        const foodGradient = this.ctx.createRadialGradient(
            this.gameState.food[0] * this.cellSize + this.cellSize / 2,
            this.gameState.food[1] * this.cellSize + this.cellSize / 2,
            0,
            this.gameState.food[0] * this.cellSize + this.cellSize / 2,
            this.gameState.food[1] * this.cellSize + this.cellSize / 2,
            this.cellSize / 2 * pulse
        );
        foodGradient.addColorStop(0, '#ff6b6b');
        foodGradient.addColorStop(0.7, '#ff8e53');
        foodGradient.addColorStop(1, '#fbbf24');
        
        this.ctx.fillStyle = foodGradient;
        this.ctx.beginPath();
        this.ctx.arc(
            this.gameState.food[0] * this.cellSize + this.cellSize / 2,
            this.gameState.food[1] * this.cellSize + this.cellSize / 2,
            this.cellSize / 3 * pulse,
            0,
            Math.PI * 2
        );
        this.ctx.fill();
        
        // Add glow effect to food
        this.ctx.shadowColor = '#ff6b6b';
        this.ctx.shadowBlur = 10 * pulse;
        this.ctx.fill();
        this.ctx.shadowBlur = 0;
    }
    
    gameLoop() {
        this.fetchGameState();
        requestAnimationFrame(() => this.gameLoop());
    }
}

// Initialize game when page loads
document.addEventListener('DOMContentLoaded', () => {
    new SnakeGameClient();
});
