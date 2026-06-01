class Game {
    constructor() {
        this.canvas = document.getElementById('gameCanvas');
        this.ctx = this.canvas.getContext('2d');
        this.width = this.canvas.width;
        this.height = this.canvas.height;
        
        this.player = null;
        this.enemies = [];
        this.bullets = [];
        this.particles = [];
        this.stars = [];
        
        this.score = 0;
        this.lives = 3;
        this.level = 1;
        this.gameState = 'menu';
        this.isPaused = false;
        this.soundEnabled = true;
        
        this.keys = {};
        this.lastTime = 0;
        this.enemySpawnTimer = 0;
        this.enemySpawnInterval = 2000;
        
        this.init();
    }
    
    init() {
        this.setupEventListeners();
        this.createStars();
        this.createPlayer();
        this.gameLoop();
    }
    
    setupEventListeners() {
        document.addEventListener('keydown', (e) => {
            this.keys[e.code] = true;
            
            if (e.code === 'Escape') {
                this.togglePause();
            }
            if (e.code === 'KeyP') {
                this.resume();
            }
        });
        
        document.addEventListener('keyup', (e) => {
            this.keys[e.code] = false;
        });
        
        document.getElementById('startBtn').addEventListener('click', () => {
            this.startGame();
        });
        
        document.getElementById('pauseBtn').addEventListener('click', () => {
            this.togglePause();
        });
        
        document.getElementById('soundBtn').addEventListener('click', () => {
            this.toggleSound();
        });
    }
    
    createStars() {
        for (let i = 0; i < 100; i++) {
            this.stars.push({
                x: Math.random() * this.width,
                y: Math.random() * this.height,
                size: Math.random() * 2 + 0.5,
                speed: Math.random() * 1 + 0.5
            });
        }
    }
    
    createPlayer() {
        this.player = new Player(this.width / 2, this.height - 100, this);
    }
    
    startGame() {
        this.gameState = 'playing';
        this.score = 0;
        this.lives = 3;
        this.level = 1;
        this.enemies = [];
        this.bullets = [];
        this.particles = [];
        this.enemySpawnInterval = 2000;
        this.createPlayer();
        this.updateUI();
        this.playSound('start');
    }
    
    togglePause() {
        if (this.gameState === 'playing') {
            this.isPaused = !this.isPaused;
        }
    }
    
    resume() {
        if (this.gameState === 'playing') {
            this.isPaused = false;
        }
    }
    
    toggleSound() {
        this.soundEnabled = !this.soundEnabled;
        document.getElementById('soundBtn').textContent = `声音: ${this.soundEnabled ? '开' : '关'}`;
    }
    
    playSound(type) {
        if (!this.soundEnabled) return;
        
        const audioContext = new (window.AudioContext || window.webkitAudioContext)();
        const oscillator = audioContext.createOscillator();
        const gainNode = audioContext.createGain();
        
        oscillator.connect(gainNode);
        gainNode.connect(audioContext.destination);
        
        switch(type) {
            case 'shoot':
                oscillator.frequency.value = 800;
                gainNode.gain.value = 0.1;
                oscillator.start();
                oscillator.stop(audioContext.currentTime + 0.05);
                break;
            case 'explosion':
                oscillator.frequency.value = 200;
                gainNode.gain.value = 0.2;
                oscillator.start();
                oscillator.stop(audioContext.currentTime + 0.1);
                break;
            case 'hit':
                oscillator.frequency.value = 300;
                gainNode.gain.value = 0.15;
                oscillator.start();
                oscillator.stop(audioContext.currentTime + 0.08);
                break;
            case 'start':
                oscillator.frequency.value = 600;
                gainNode.gain.value = 0.1;
                oscillator.frequency.exponentialRampToValueAtTime(1200, audioContext.currentTime + 0.2);
                oscillator.start();
                oscillator.stop(audioContext.currentTime + 0.2);
                break;
            case 'gameOver':
                oscillator.frequency.value = 400;
                gainNode.gain.value = 0.2;
                oscillator.frequency.exponentialRampToValueAtTime(100, audioContext.currentTime + 0.5);
                oscillator.start();
                oscillator.stop(audioContext.currentTime + 0.5);
                break;
        }
    }
    
    update(deltaTime) {
        if (this.gameState !== 'playing' || this.isPaused) return;
        
        this.updateStars();
        this.player.update(deltaTime);
        
        this.updateBullets(deltaTime);
        this.updateEnemies(deltaTime);
        this.updateParticles(deltaTime);
        
        this.spawnEnemies(deltaTime);
        this.checkCollisions();
        this.checkLevelUp();
        
        if (this.lives <= 0) {
            this.gameOver();
        }
    }
    
    updateStars() {
        this.stars.forEach(star => {
            star.y += star.speed;
            if (star.y > this.height) {
                star.y = -10;
                star.x = Math.random() * this.width;
            }
        });
    }
    
    updateBullets(deltaTime) {
        this.bullets = this.bullets.filter(bullet => {
            bullet.update(deltaTime);
            return bullet.y > -10;
        });
    }
    
    updateEnemies(deltaTime) {
        this.enemies = this.enemies.filter(enemy => {
            enemy.update(deltaTime);
            return enemy.y < this.height + 50 && enemy.hp > 0;
        });
    }
    
    updateParticles(deltaTime) {
        this.particles = this.particles.filter(particle => {
            particle.update(deltaTime);
            return particle.life > 0;
        });
    }
    
    spawnEnemies(deltaTime) {
        this.enemySpawnTimer += deltaTime;
        
        if (this.enemySpawnTimer >= this.enemySpawnInterval) {
            const x = Math.random() * (this.width - 40) + 20;
            const enemyType = Math.random() < 0.7 ? 'basic' : 'fast';
            this.enemies.push(new Enemy(x, -50, enemyType, this));
            this.enemySpawnTimer = 0;
        }
    }
    
    checkCollisions() {
        this.bullets.forEach((bullet, bulletIndex) => {
            this.enemies.forEach((enemy, enemyIndex) => {
                if (this.checkCollision(bullet, enemy)) {
                    enemy.hp -= bullet.damage;
                    this.createParticles(bullet.x, bullet.y, 'hit');
                    this.playSound('hit');
                    
                    if (enemy.hp <= 0) {
                        this.score += enemy.points;
                        this.createParticles(enemy.x, enemy.y, 'explosion');
                        this.playSound('explosion');
                        this.enemies.splice(enemyIndex, 1);
                        this.updateUI();
                    }
                    
                    this.bullets.splice(bulletIndex, 1);
                }
            });
        });
        
        this.enemies.forEach((enemy, index) => {
            if (this.checkCollision(this.player, enemy)) {
                this.lives--;
                this.createParticles(this.player.x, this.player.y, 'explosion');
                this.playSound('hit');
                this.enemies.splice(index, 1);
                this.updateUI();
                
                if (this.lives > 0) {
                    this.player.respawn();
                }
            }
        });
    }
    
    checkCollision(obj1, obj2) {
        return obj1.x < obj2.x + obj2.width &&
               obj1.x + obj1.width > obj2.x &&
               obj1.y < obj2.y + obj2.height &&
               obj1.y + obj1.height > obj2.y;
    }
    
    createParticles(x, y, type) {
        const count = type === 'explosion' ? 20 : 10;
        const color = type === 'explosion' ? '#FF6B35' : '#FFD700';
        
        for (let i = 0; i < count; i++) {
            this.particles.push(new Particle(x, y, color));
        }
    }
    
    checkLevelUp() {
        const newLevel = Math.floor(this.score / 100) + 1;
        if (newLevel > this.level) {
            this.level = newLevel;
            this.enemySpawnInterval = Math.max(500, 2000 - (this.level - 1) * 200);
            this.updateUI();
        }
    }
    
    gameOver() {
        this.gameState = 'gameover';
        this.playSound('gameOver');
    }
    
    updateUI() {
        document.getElementById('score').textContent = this.score;
        document.getElementById('lives').textContent = this.lives;
        document.getElementById('level').textContent = this.level;
    }
    
    render() {
        this.ctx.fillStyle = '#000033';
        this.ctx.fillRect(0, 0, this.width, this.height);
        
        this.renderStars();
        
        if (this.gameState === 'playing') {
            this.player.render(this.ctx);
            this.enemies.forEach(enemy => enemy.render(this.ctx));
            this.bullets.forEach(bullet => bullet.render(this.ctx));
            this.particles.forEach(particle => particle.render(this.ctx));
            
            if (this.isPaused) {
                this.renderPauseScreen();
            }
        } else if (this.gameState === 'menu') {
            this.renderMenuScreen();
        } else if (this.gameState === 'gameover') {
            this.renderGameOverScreen();
        }
    }
    
    renderStars() {
        this.ctx.fillStyle = 'white';
        this.stars.forEach(star => {
            this.ctx.globalAlpha = star.size / 2.5;
            this.ctx.fillRect(star.x, star.y, star.size, star.size);
        });
        this.ctx.globalAlpha = 1;
    }
    
    renderMenuScreen() {
        this.ctx.fillStyle = '#FFD700';
        this.ctx.font = 'bold 48px Arial';
        this.ctx.textAlign = 'center';
        this.ctx.fillText('飞机大战', this.width / 2, this.height / 2 - 50);
        
        this.ctx.font = '24px Arial';
        this.ctx.fillStyle = 'white';
        this.ctx.fillText('点击 "开始游戏" 开始', this.width / 2, this.height / 2 + 20);
        
        this.ctx.font = '18px Arial';
        this.ctx.fillText('使用方向键控制移动，空格键发射子弹', this.width / 2, this.height / 2 + 60);
    }
    
    renderPauseScreen() {
        this.ctx.fillStyle = 'rgba(0, 0, 0, 0.5)';
        this.ctx.fillRect(0, 0, this.width, this.height);
        
        this.ctx.fillStyle = '#FFD700';
        this.ctx.font = 'bold 36px Arial';
        this.ctx.textAlign = 'center';
        this.ctx.fillText('游戏暂停', this.width / 2, this.height / 2);
        
        this.ctx.font = '20px Arial';
        this.ctx.fillStyle = 'white';
        this.ctx.fillText('按 P 键继续', this.width / 2, this.height / 2 + 40);
    }
    
    renderGameOverScreen() {
        this.ctx.fillStyle = '#FF6B35';
        this.ctx.font = 'bold 48px Arial';
        this.ctx.textAlign = 'center';
        this.ctx.fillText('游戏结束', this.width / 2, this.height / 2 - 50);
        
        this.ctx.fillStyle = 'white';
        this.ctx.font = '24px Arial';
        this.ctx.fillText(`最终得分: ${this.score}`, this.width / 2, this.height / 2);
        this.ctx.fillText(`达到等级: ${this.level}`, this.width / 2, this.height / 2 + 40);
        
        this.ctx.font = '18px Arial';
        this.ctx.fillText('点击 "开始游戏" 重新开始', this.width / 2, this.height / 2 + 80);
    }
    
    gameLoop(currentTime = 0) {
        const deltaTime = currentTime - this.lastTime;
        this.lastTime = currentTime;
        
        this.update(deltaTime);
        this.render();
        
        requestAnimationFrame((time) => this.gameLoop(time));
    }
}

window.addEventListener('DOMContentLoaded', () => {
    new Game();
});
