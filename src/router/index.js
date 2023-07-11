import { createRouter, createWebHistory } from 'vue-router';

const routes = [
    {
        path: '/',
        redirect: '/home'
    },
    {
        path: '/home',
        name: 'home',
        component: () => import('../views/Home.vue')
    },
    {
        path: '/baseStrategy',
        name: 'baseStrategy',
        component: () => import('../views/BaseStrategy.vue')
    },
];
const router = createRouter({
    history: createWebHistory(import.meta.env.BASE_URL),
    routes
})

const originalPush = router.push
router.push = function push(location) {
    return originalPush.call(this, location).catch(err => err)
}

export default router