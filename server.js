const express = require('express');
const cors = require('cors');
const ytdl = require('@distube/ytdl-core');
const youtubedl = require('youtube-dl-exec');

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
        
        const subprocess = youtubedl.exec(url, {
            output: '-', // Pipe output to stdout
            format: formatOption,
            noCheckCertificates: true,
            noWarnings: true,
            addHeader: [
                'referer:youtube.com',
                'user-agent:Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            ]
        });

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
