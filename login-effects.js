/**
 * Login Screen Water Ripple Effect 
 * AD
 */

(function() {
  'use strict';

  const isMobile = () => {
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ||
           window.innerWidth < 768;
  };

  class WaterEffect {
    constructor() {
      this.canvas = document.createElement('canvas');
      this.canvas.id = 'water-effect-canvas';
      this.canvas.style.position = 'fixed';
      this.canvas.style.top = '0';
      this.canvas.style.left = '0';
      this.canvas.style.pointerEvents = 'none';
      this.canvas.style.zIndex = '1';
      document.body.insertBefore(this.canvas, document.body.firstChild);

      this.ctx = this.canvas.getContext('2d');
      this.isMobile = isMobile();
      this.particles = [];
      this.waves = [];
      this.lastX = 0;
      this.lastY = 0;
      this.velocityX = 0;
      this.velocityY = 0;

      this.resize();
      window.addEventListener('resize', () => this.resize());

      if (this.isMobile) {
        this.setupMobileEvents();
      } else {
        this.setupDesktopEvents();
      }

      this.animate();
    }

    resize() {
      this.canvas.width = window.innerWidth;
      this.canvas.height = window.innerHeight;
    }

    setupDesktopEvents() {
      document.addEventListener('mousemove', (e) => {
        this.velocityX = e.clientX - this.lastX;
        this.velocityY = e.clientY - this.lastY;
        this.lastX = e.clientX;
        this.lastY = e.clientY;

        // Create subtle ripples as cursor moves
        if (Math.abs(this.velocityX) > 0.5 || Math.abs(this.velocityY) > 0.5) {
          this.createRipple(e.clientX, e.clientY, 2);
        }
      });

      document.addEventListener('click', (e) => {
        this.createWaterdrop(e.clientX, e.clientY);
      });
    }

    setupMobileEvents() {
      document.addEventListener('touchmove', (e) => {
        const touch = e.touches[0];
        if (touch) {
          this.createRipple(touch.clientX, touch.clientY, 2);
        }
      }, { passive: true });

      document.addEventListener('touchend', (e) => {
        const touch = e.changedTouches[0];
        if (touch) {
          this.createWaterdrop(touch.clientX, touch.clientY);
        }
      }, { passive: true });
    }

    createRipple(x, y, radius) {
      this.waves.push({
        x,
        y,
        radius,
        maxRadius: 80,
        opacity: 0.6,
        speed: 2
      });
    }

    createWaterdrop(x, y) {
      // Create a larger, more visible drop effect
      this.waves.push({
        x,
        y,
        radius: 0,
        maxRadius: 100,
        opacity: 0.8,
        speed: 3,
        isDroplet: true
      });

      // Particle burst for visual effect
      for (let i = 0; i < 6; i++) {
        const angle = (Math.PI * 2 * i) / 6;
        const velocity = 2 + Math.random() * 2;
        this.particles.push({
          x,
          y,
          vx: Math.cos(angle) * velocity,
          vy: Math.sin(angle) * velocity,
          life: 1,
          size: 2 + Math.random() * 2
        });
      }
    }

    animate() {
      // Clear with slight trail effect for motion blur
      this.ctx.fillStyle = 'rgba(255, 255, 255, 0.02)';
      this.ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);

      // Update and draw waves
      this.waves = this.waves.filter(wave => wave.radius < wave.maxRadius);
      for (const wave of this.waves) {
        wave.radius += wave.speed;
        wave.opacity = 0.6 * (1 - wave.radius / wave.maxRadius);

        this.ctx.strokeStyle = `rgba(147, 197, 253, ${wave.opacity})`;
        this.ctx.lineWidth = wave.isDroplet ? 3 : 1.5;
        this.ctx.beginPath();
        this.ctx.arc(wave.x, wave.y, wave.radius, 0, Math.PI * 2);
        this.ctx.stroke();
      }

      // Update and draw particles
      this.particles = this.particles.filter(p => p.life > 0);
      for (const particle of this.particles) {
        particle.x += particle.vx;
        particle.y += particle.vy;
        particle.vy += 0.1; // Gravity
        particle.life -= 0.02;
        particle.vx *= 0.98; // Friction

        this.ctx.fillStyle = `rgba(147, 197, 253, ${particle.life * 0.5})`;
        this.ctx.beginPath();
        this.ctx.arc(particle.x, particle.y, particle.size * particle.life, 0, Math.PI * 2);
        this.ctx.fill();
      }

      requestAnimationFrame(() => this.animate());
    }

    destroy() {
      if (this.canvas && this.canvas.parentNode) {
        this.canvas.parentNode.removeChild(this.canvas);
      }
    }
  }

  // Initialize when vault screen is shown
  function initWaterEffect() {
    const vaultModal = document.getElementById('vaultModal');
    if (!vaultModal) return;

    let effectInstance = null;

    // Watch for vault modal visibility
    const observer = new MutationObserver((mutations) => {
      const isVisible = vaultModal.classList.contains('show');

      if (isVisible && !effectInstance) {
        effectInstance = new WaterEffect();
      } else if (!isVisible && effectInstance) {
        effectInstance.destroy();
        effectInstance = null;
      }
    });

    observer.observe(vaultModal, { attributes: true, attributeFilter: ['class'] });

    // Initial check
    if (vaultModal.classList.contains('show')) {
      effectInstance = new WaterEffect();
    }
  }

  // Start when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initWaterEffect);
  } else {
    initWaterEffect();
  }
})();
