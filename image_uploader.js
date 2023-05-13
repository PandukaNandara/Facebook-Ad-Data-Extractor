
import got from "got"


const API_KEY = '6d207e02198a847aa98d0a2a901485a5';


export const uploadImage = async (imageUrl) => {
    const searchParams = new URLSearchParams([['key', API_KEY], ['format', 'txt'], ['source', imageUrl]]);
    const rst = await got.post('https://freeimage.host/api/1/upload', {
        searchParams
    })
    return rst?.body;
}