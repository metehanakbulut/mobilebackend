const express = require('express');
const cors = require('cors');
const ytdl = require('@distube/ytdl-core');

const app = express();
app.use(cors());
app.use(express.json());

app.get('/api/youtube-download', async (req, res) => {
    try {
        const { url, itag } = req.query;
        if (!ytdl.validateURL(url)) return res.status(400).json({ error: 'Geçersiz URL' });

        const isMp3 = itag === 'bestaudio/best';
        res.header('Content-Disposition', `attachment; filename="indirme_${Date.now()}.${isMp3 ? 'mp3' : 'mp4'}"`);
        
        if (isMp3) {
            ytdl(url, { filter: 'audioonly', quality: 'highestaudio' }).pipe(res);
        } else {
            ytdl(url, { filter: 'audioandvideo', quality: 'highest' }).pipe(res);
        }
    } catch (error) {
        if (!res.headersSent) res.status(500).json({ error: 'Hata oluştu' });
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Mobile Backend Server baslatildi! Port: ${PORT}`);
});
