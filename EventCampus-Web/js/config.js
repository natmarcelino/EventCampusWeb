// js/config.js — endereço da API Spring Boot
//
// Abrindo pelo Live Server (localhost / 127.0.0.1) → usa a API do seu computador.
// Abrindo pelo site publicado (Render / seu domínio)  → usa a API online.
//
// ATENÇÃO: depois de criar a API no Render, troque o endereço abaixo pelo dela.
const API_URL_ONLINE = 'https://eventcampus-api.onrender.com';
const API_URL_LOCAL  = 'http://localhost:8080';

const API_URL = ['localhost', '127.0.0.1', ''].includes(location.hostname)
  ? API_URL_LOCAL
  : API_URL_ONLINE;

// O servidor grátis do Render "dorme" depois de 15 min sem uso.
// Esta chamada acorda a API assim que a página abre, para o login não demorar.
fetch(API_URL + '/api/status').catch(() => {});