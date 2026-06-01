class Player {
    constructor(x, y, game) {
        this.x = x;
        this.y = y;
        this.width = 40;
        this.height = 40;
        this.speed = 5;
        this.game = game;
        this.shootCooldown = 0;
        this.shootInterval = 200;
        this.isInvulnerable = false;
        this.invulnerableTimer = 0;
    }
    
    update(deltaTime) {
        if (this.game.keys['ArrowLeft'] || this.game.keys['KeyA']) {
            this.x = Math.max(0, this.x - this.speed);
        }
        if (this.game.keys['ArrowRight'] || this.game.keys['KeyD']) {
            this.x = Math.min(this.game.width - this.width, this.x + this.speed);
        }
        if (this.game.keys['ArrowUp'] || this.game.keys['KeyW']) {
            this.y = Math.max(0, this.y - this.speed);
        }
        if (this.game.keys['ArrowDown'] || this.game.keys['KeyS']) {
            this.y = Math.min(this.game.height - this.height, this.y + this.speed);
        }
        
        if (this.game.keys['Space']) {
            this.shoot();
        }
        
        if (this.shootCooldown > 0) {
            this.shootCooldown -= deltaTime;
        }
        
        if (this.isInvulnerable) {
            this.invulnerableTimer -= deltaTime;
            if (this.invulnerableTimer <= 0) {
                this.isInvulnerable = false;
            }
        }
    }
    
    shoot() {
        if (this.shootCooldown <= 0) {
            this.game.bullets.push(new Bullet(this.x + this.width / 2 - 2, this.y, -10, 'player'));
            this.game.playSound('shoot');
            this.shootCooldown = this.shootInterval;
        }
    }
    
    respawn() {
        this.x = this.game.width / 2 - this.width / 2;
        this.y = this.game.height - 100;
        this.isInvulnerable = true;
        this.invulnerableTimer = 2000;
    }
    
    render(ctx) {
        if (this.isInvulnerable && Math.floor(Date.now() / 100) % 2 === 0) {
            ctx.globalAlpha = 0.5;
        }
        
        ctx.fillStyle = '#4CAF50';
        ctx.fillRect(this.x, this.y, this.width, this.height);
        
        ctx.fillStyle = '#81C784';
        ctx.fillRect(this.x + 5, this.y + 5, this.width - 10, 10);
        
        ctx.fillStyle = '#FFD700';
        ctx.fillRect(this.x + this.width / 2 - 3, this.y - 5, 6, 10);
        
        ctx.fillStyle = '#FFA500';
        ctx.fillRect(this.x + 8, this.y + 25, 8, 15);
        ctx.fillRect(this.x + this.width - 16, this.y + 25, 8, 15);
        
        ctx.globalAlpha = 1;
    }
}

class Enemy {
    constructor(x, y, type, game) {
        this.x = x;
        this.y = y;
        this.type = type;
        this.game = game;
        
        if (type === 'basic') {
            this.width = 30;
            this.height = 30;
            this.speed = 2;
            this.hp = 1;
            this.points = 10;
            this.color = '#FF6B35';
        } else if (type === 'fast') {
            this.width = 25;
            this.height = 25;
            this.speed = 4;
            this.hp = 1;
            this.points = 20;
            this.color = '#E91E63';
        }
    }
    
    update(deltaTime) {
        this.y += this.speed;
        
        if (this.type === 'fast') {
            this.x += Math.sin(this.y * 0.02) * 2;
        }
    }
    
    render(ctx) {
        ctx.fillStyle = this.color;
        ctx.fillRect(this.x, this.y, this.width, this.height);
        
        ctx.fillStyle = '#FFD700';
        ctx.fillRect(this.x + 5, this.y + 5, this.width - 10, 8);
        
        ctx.fillStyle = '#8B4513';
        ctx.fillRect(this.x + 8, this.y + this.height - 8, 6, 8);
        ctx.fillRect(this.x + this.width - 14, this.y + this.height - 8, 6, 8);
    }
}

class Bullet {
    constructor(x, y, speed, owner) {
        this.x = x;
        this.y = y;
        this.width = 4;
        this.height = 10;
        this.speed = speed;
        this.owner = owner;
        this.damage = 1;
    }
    
    update(deltaTime) {
        this.y += this.speed;
    }
    
    render(ctx) {
        ctx.fillStyle = this.owner === 'player' ? '#FFD700' : '#FF6B35';
        ctx.fillRect(this.x, this.y, this.width, this.height);
        
        ctx.fillStyle = 'white';
        ctx.fillRect(this.x + 1, this.y + 2, 2, 6);
    }
}

class Particle {
    constructor(x, y, color) {
        this.x = x;
        this.y = y;
        this.vx = (Math.random() - 0.5) * 8;
        this.vy = (Math.random() - 0.5) * 8;
        this.color = color;
        this.life = 1.0;
        this.decay = 0.02;
        this.size = Math.random() * 3 + 2;
    }
    
    update(deltaTime) {
        this.x += this.vx;
        this.y += this.vy;
        this.life -= this.decay;
        this.vx *= 0.98;
        this.vy *= 0.98;
    }
    
    render(ctx) {
        ctx.globalAlpha = this.life;
        ctx.fillStyle = this.color;
        ctx.fillRect(this.x, this.y, this.size, this.size);
        ctx.globalAlpha = 1;
    }
}
