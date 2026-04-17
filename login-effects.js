/**
 * Login Screen Water Ripple Effects
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
      this.canvas.style.position = 'absolute';
      this.canvas.style.top = '0';
      this.canvas.style.left = '0';
      this.canvas.style.width = '100%';
      this.canvas.style.height = '100%';
      this.canvas.style.background = 'transparent';
      this.canvas.style.pointerEvents = 'none';
      this.canvas.style.zIndex = '1';

      const vaultScreen = document.getElementById('vaultScreen');
      if (vaultScreen) {
        // Keep effect confined to login screen only and behind the login panel.
        if (!vaultScreen.style.position) vaultScreen.style.position = 'relative';
        vaultScreen.style.overflow = 'hidden';
        vaultScreen.insertBefore(this.canvas, vaultScreen.firstChild);
      } else {
        // Safe fallback: if login container is missing, do not show any effect.
        throw new Error('vaultScreen not found for water effect');
      }

      this.ctx = this.canvas.getContext('2d');
      this.isMobile = isMobile();
      this.particles = [];
      this.waves = [];
      this.lastX = 0;
      this.lastY = 0;
      this.velocityX = 0;
      this.velocityY = 0;
      this.lastRippleAt = 0;
      this.hasPointer = false;
      this._onMouseMove = null;
      this._onClick = null;
      this._onTouchMove = null;
      this._onTouchEnd = null;

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
      const host = document.getElementById('vaultScreen');
      if (!host) return;
      const rect = host.getBoundingClientRect();
      const dpr = Math.max(1, window.devicePixelRatio || 1);
      this.canvas.width = Math.max(1, Math.floor(rect.width * dpr));
      this.canvas.height = Math.max(1, Math.floor(rect.height * dpr));
      this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    }

    setupDesktopEvents() {
      this._onMouseMove = (e) => {
        if (!shouldRenderEffect()) return;
        if (isInsideLoginPanel(e.target)) return;
        this.hasPointer = true;
        this.velocityX = e.clientX - this.lastX;
        this.velocityY = e.clientY - this.lastY;
        this.lastX = e.clientX;
        this.lastY = e.clientY;

        // Gentle trail: only add ripples when cursor movement is meaningful and throttled.
        const now = Date.now();
        const speed = Math.abs(this.velocityX) + Math.abs(this.velocityY);
        if (speed > 4 && (now - this.lastRippleAt) > 100) {
          this.createRipple(e.clientX, e.clientY, 1.2);
          this.lastRippleAt = now;
        }
      };
      document.addEventListener('mousemove', this._onMouseMove);

      this._onClick = (e) => {
        if (!shouldRenderEffect()) return;
        if (isInsideLoginPanel(e.target)) return;
        this.createWaterdrop(e.clientX, e.clientY, 1);
      };
      document.addEventListener('click', this._onClick);
    }

    setupMobileEvents() {
      this._onTouchMove = (e) => {
        if (!shouldRenderEffect()) return;
        if (isInsideLoginPanel(e.target)) return;
        const touch = e.touches[0];
        if (touch) {
          this.hasPointer = true;
          const now = Date.now();
          if ((now - this.lastRippleAt) > 120) {
            this.createRipple(touch.clientX, touch.clientY, 1.1);
            this.lastRippleAt = now;
          }
        }
      };
      document.addEventListener('touchmove', this._onTouchMove, { passive: true });

      this._onTouchEnd = (e) => {
        if (!shouldRenderEffect()) return;
        if (isInsideLoginPanel(e.target)) return;
        const touch = e.changedTouches[0];
        if (touch) {
          this.createWaterdrop(touch.clientX, touch.clientY, 0.9);
        }
      };
      document.addEventListener('touchend', this._onTouchEnd, { passive: true });
    }

    createRipple(x, y, radius) {
      this.waves.push({
        x,
        y,
        radius,
        maxRadius: 58,
        opacity: 0.22,
        speed: 0.48,
        lineWidth: 1.5
      });
    }

    createWaterdrop(x, y, scale) {
      // Small, soft concentric drop rings.
      this.waves.push({
        x,
        y,
        radius: 0,
        maxRadius: 72 * scale,
        opacity: 0.28,
        speed: 0.62,
        lineWidth: 1.8,
        isDroplet: true,
        delay: 0
      });
      this.waves.push({
        x,
        y,
        radius: 0,
        maxRadius: 98 * scale,
        opacity: 0.21,
        speed: 0.56,
        lineWidth: 1.3,
        isDroplet: true,
        delay: 110
      });
    }

    animate() {
      try {
        // Keep canvas fully transparent so the login background remains unchanged.
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);

        // Update and draw waves
        this.waves = this.waves.filter(wave => {
          if (wave.delay && wave.delay > 0) return true;
          return wave.radius < wave.maxRadius;
        });
        for (const wave of this.waves) {
          if (wave.delay && wave.delay > 0) {
            wave.delay -= 16;
            continue;
          }
          wave.radius += wave.speed;
          wave.opacity = (wave.isDroplet ? 0.4 : 0.3) * (1 - wave.radius / wave.maxRadius);
          if (wave.opacity <= 0.001) continue;

          this.ctx.shadowBlur = wave.isDroplet ? 12 : 8;
          this.ctx.shadowColor = `rgba(8, 47, 73, ${Math.min(0.24, wave.opacity)})`;
          this.ctx.strokeStyle = `rgba(3, 105, 161, ${Math.min(0.6, wave.opacity * 1.45)})`;
          this.ctx.lineWidth = (wave.lineWidth || (wave.isDroplet ? 1 : 0.9)) + 0.2;
          this.ctx.beginPath();
          this.ctx.arc(wave.x, wave.y, wave.radius, 0, Math.PI * 2);
          this.ctx.stroke();

          this.ctx.shadowBlur = 0;
          this.ctx.strokeStyle = `rgba(186, 230, 253, ${Math.min(0.4, wave.opacity * 0.9)})`;
          this.ctx.lineWidth = Math.max(0.8, (wave.lineWidth || 1) * 0.52);
          this.ctx.beginPath();
          this.ctx.arc(wave.x, wave.y, Math.max(0, wave.radius - 0.6), 0, Math.PI * 2);
          this.ctx.stroke();
        }
      } catch (_) {
        // Never allow effect errors to stop the animation loop.
      }

      requestAnimationFrame(() => this.animate());
    }

    destroy() {
      if (this._onMouseMove) document.removeEventListener('mousemove', this._onMouseMove);
      if (this._onClick) document.removeEventListener('click', this._onClick);
      if (this._onTouchMove) document.removeEventListener('touchmove', this._onTouchMove);
      if (this._onTouchEnd) document.removeEventListener('touchend', this._onTouchEnd);
      if (this.canvas && this.canvas.parentNode) {
        this.canvas.parentNode.removeChild(this.canvas);
      }
    }
  }

  function isInsideLoginPanel(target) {
    if (!target || !target.closest) return false;
    return !!target.closest('.vault-box');
  }

  function shouldRenderEffect() {
    const vaultScreen = document.getElementById('vaultScreen');
    if (!vaultScreen) return false;
    const appShell = document.getElementById('appShell');
    const screenVisible = window.getComputedStyle(vaultScreen).display !== 'none';
    const appUnlocked = !!(appShell && appShell.classList.contains('unlocked'));
    return screenVisible && !appUnlocked;
  }

  // Initialize when vault screen is shown
  function initWaterEffect() {
    const vaultScreen = document.getElementById('vaultScreen');
    if (!vaultScreen) return;

    let effectInstance = null;

    function syncEffect() {
      const isVisible = shouldRenderEffect();
      if (isVisible && !effectInstance) {
        effectInstance = new WaterEffect();
      } else if (!isVisible && effectInstance) {
        effectInstance.destroy();
        effectInstance = null;
      }
    }

    // Watch for vault screen visibility changes
    const observer = new MutationObserver(() => {
      syncEffect();
    });

    observer.observe(vaultScreen, { attributes: true, attributeFilter: ['style', 'class'] });
    document.addEventListener('handymgr:login-screen-visible', syncEffect);

    // Initial check
    syncEffect();
  }

  // Start when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initWaterEffect);
  } else {
    initWaterEffect();
  }
})();
