import { mount } from "svelte";
import App from "./App.svelte";
import { i18n } from "./i18n/index.svelte";
import "./styles.css";

document.documentElement.lang = i18n.lang;
mount(App, { target: document.getElementById("app")! });
