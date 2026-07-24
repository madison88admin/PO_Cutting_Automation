/** @type {import('next').NextConfig} */
const uploadAllowedOrigin = (process.env.ALLOWED_UPLOAD_ORIGINS || 'https://m88-po-cutting.netlify.app')
    .split(',')[0]
    .trim()
    .replace(/\/+$/, '');

const nextConfig = {
    async headers() {
        return [
            {
                source: '/api/:path*',
                headers: [
                    { key: 'Access-Control-Allow-Origin', value: uploadAllowedOrigin },
                    { key: 'Access-Control-Allow-Methods', value: 'POST, OPTIONS' },
                    { key: 'Access-Control-Allow-Headers', value: 'Content-Type' },
                    { key: 'Vary', value: 'Origin' },
                ],
            },
        ];
    },
};

export default nextConfig;
