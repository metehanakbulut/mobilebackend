const express = require('express');
const cors = require('cors');
const ytdl = require('@distube/ytdl-core');
const { spawn } = require('child_process');
const fs = require('fs');
const https = require('https');

const YTDLP_PATH = './yt-dlp';

// Startup Check: Download yt-dlp if it doesn't exist
if (!fs.existsSync(YTDLP_PATH)) {
    console.log("yt-dlp binary not found, downloading from GitHub...");
    const file = fs.createWriteStream(YTDLP_PATH);
    https.get("https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp", (response) => {
        if (response.statusCode === 302 || response.statusCode === 301) {
            https.get(response.headers.location, (res2) => {
                res2.pipe(file);
                file.on("finish", () => {
                    file.close();
                    fs.chmodSync(YTDLP_PATH, '755');
                    console.log("yt-dlp downloaded successfully!");
                });
            });
        } else {
            response.pipe(file);
            file.on("finish", () => {
                file.close();
                fs.chmodSync(YTDLP_PATH, '755');
                console.log("yt-dlp downloaded successfully!");
            });
        }
    }).on('error', (err) => {
        console.error("Failed to download yt-dlp:", err.message);
    });
}

const app = express();
app.use(cors());
app.use(express.json());

// 1. YouTube Bilgilerini Getirme
app.get('/api/youtube-info', async (req, res) => {
    try {
        const { url } = req.query;
        if (!ytdl.validateURL(url)) {
            return res.status(400).json({ error: 'Geçersiz YouTube URL\'si' });
        }
        
        const info = await ytdl.getInfo(url);
        const title = info.videoDetails.title;
        const thumbnail = info.videoDetails.thumbnails.length > 0 
            ? info.videoDetails.thumbnails[info.videoDetails.thumbnails.length - 1].url 
            : null;

        res.json({ title, thumbnail });
    } catch (error) {
        console.error("Hata (youtube-info):", error.message);
        res.status(500).json({ error: 'Video bilgileri alınamadı.' });
    }
});

// 2. YouTube İndirme
app.get('/api/youtube-download', async (req, res) => {
    try {
        const { url, itag } = req.query;
        if (!ytdl.validateURL(url)) {
            return res.status(400).json({ error: 'Geçersiz YouTube URL\'si' });
        }

        const isMp3 = itag === 'bestaudio/best';
        const fileExt = isMp3 ? 'mp3' : 'mp4';
        
        res.header('Content-Disposition', `attachment; filename="indirme_${Date.now()}.${fileExt}"`);
        
        const formatOption = isMp3 ? 'bestaudio' : 'best';
        
        const args = [
            url,
            '-o', '-',
            '-f', formatOption,
            '--no-check-certificates',
            '--no-warnings',
            '--add-header', 'referer:youtube.com',
            '--add-header', 'user-agent:Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        ];

        const subprocess = spawn('./yt-dlp', args);

        subprocess.stdout.pipe(res);

        subprocess.on('error', (err) => {
            console.error('YTDLP Error:', err.message);
            if (!res.headersSent) res.status(500).end();
        });

    } catch (error) {
        console.error("Hata (youtube-download):", error.message);
        if (!res.headersSent) {
            res.status(500).json({ error: 'İndirme sırasında bir hata oluştu.' });
        }
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Mobile Backend Server başlatıldı! (Port: ${PORT})`);
    console.log(`Telefondan erişim için bilgisayarınızın yerel IP adresini kullanın.`);
});
