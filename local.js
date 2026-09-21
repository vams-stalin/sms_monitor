(function () {
    const canvas = document.getElementById('bg-canvas');
    const ctx    = canvas.getContext('2d');
    let W, H, rings = [], floats = [];

    function resize() {
        W = canvas.width  = window.innerWidth;
        H = canvas.height = window.innerHeight;
    }

    /*  Slow-drifting soft rings */
    function Ring(i) {
        this.reset(i);
    }
    Ring.prototype.reset = function (i) {
        this.cx    = W * 0.5 + (Math.random() - 0.5) * W * 0.6;
        this.cy    = H * 0.5 + (Math.random() - 0.5) * H * 0.6;
        this.r     = 80 + Math.random() * 180;
        this.maxR  = this.r + 60 + Math.random() * 80;
        this.alpha = 0;
        this.growing = true;
        this.speed = 0.18 + Math.random() * 0.14;
        this.hue   = Math.random() < 0.5 ? '200,190,220' : '180,210,200';
        this.delay = i * 900;
        this.born  = Date.now() + this.delay;
    };
    Ring.prototype.draw = function () {
        const now = Date.now();
        if (now < this.born) return;
        if (this.growing) {
            this.r     += this.speed;
            this.alpha  = Math.min(this.alpha + 0.004, 0.13);
            if (this.r >= this.maxR) this.growing = false;
        } else {
            this.alpha -= 0.003;
            if (this.alpha <= 0) { this.reset(0); return; }
        }
        ctx.beginPath();
        ctx.arc(this.cx, this.cy, this.r, 0, Math.PI * 2);
        ctx.strokeStyle = `rgba(${this.hue},${this.alpha})`;
        ctx.lineWidth   = 1.2;
        ctx.stroke();
    };

    /* Tiny floating specks */
    function Speck() {
        this.x     = Math.random() * W;
        this.y     = Math.random() * H;
        this.size  = Math.random() * 3 + 1;
        this.vx    = (Math.random() - 0.5) * 0.18;
        this.vy    = -(Math.random() * 0.18 + 0.06);
        this.alpha = Math.random() * 0.18 + 0.06;
        this.color = Math.random() < 0.5 ? '37,99,168' : '100,130,160';
    }
    Speck.prototype.step = function () {
        this.x += this.vx;
        this.y += this.vy;
        this.alpha -= 0.0003;
        if (this.y < -10 || this.alpha <= 0) {
            this.y     = H + 10;
            this.x     = Math.random() * W;
            this.alpha = Math.random() * 0.18 + 0.06;
        }
    };
    Speck.prototype.draw = function () {
        ctx.beginPath();
        ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
        ctx.fillStyle = `rgba(${this.color},${this.alpha})`;
        ctx.fill();
    };

    function init() {
        rings  = [];
        floats = [];
        for (let i = 0; i < 6; i++)  rings.push(new Ring(i));
        for (let i = 0; i < 38; i++) floats.push(new Speck());
    }

    function frame() {
        ctx.clearRect(0, 0, W, H);
        rings.forEach(r  => r.draw());
        floats.forEach(s => { s.step(); s.draw(); });
        requestAnimationFrame(frame);
    }

    window.addEventListener('resize', () => { resize(); init(); });
    resize();
    init();
    frame();
})();