'use strict';
const menu=document.querySelector('.menu-toggle');
const navigation=document.getElementById('navigation');
function closeMenu(){navigation.classList.remove('open');menu.setAttribute('aria-expanded','false');menu.setAttribute('aria-label','Open navigation');}
menu.addEventListener('click',()=>{const open=menu.getAttribute('aria-expanded')!=='true';navigation.classList.toggle('open',open);menu.setAttribute('aria-expanded',String(open));menu.setAttribute('aria-label',open?'Close navigation':'Open navigation');});
navigation.querySelectorAll('a').forEach(a=>a.addEventListener('click',closeMenu));
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&navigation.classList.contains('open')){closeMenu();menu.focus();}});
document.addEventListener('click',e=>{if(!e.target.closest('.header'))closeMenu();});
function updateTime(){document.getElementById('local-time').textContent=new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Amman',hour:'2-digit',minute:'2-digit',hour12:false}).format(new Date());}
updateTime();setInterval(updateTime,60000);
const performanceButtons=[...document.querySelectorAll('[data-performance]')];
performanceButtons.forEach(button=>button.addEventListener('click',()=>{const after=button.dataset.performance==='after';performanceButtons.forEach(b=>b.setAttribute('aria-pressed',String(b===button)));document.getElementById('order-value').textContent=after?'120,000':'15,000';document.getElementById('latency-value').innerHTML=after?'&lt;200<span>ms</span>':'3<span>s</span>';document.getElementById('performance-fill').style.width=after?'100%':'12.5%';document.getElementById('performance-note').textContent=after?'8× the capacity. A fraction of the wait.':'The legacy platform, before modernization.';}));
let scrollPending=false;
function updateProgress(){const max=document.documentElement.scrollHeight-innerHeight;document.querySelector('.reading-progress').style.width=`${max>0?scrollY/max*100:0}%`;scrollPending=false;}
addEventListener('scroll',()=>{if(!scrollPending){requestAnimationFrame(updateProgress);scrollPending=true;}},{passive:true});addEventListener('resize',updateProgress);updateProgress();
const filterButtons=[...document.querySelectorAll('[data-filter]')];
const expertiseCards=[...document.querySelectorAll('[data-category]')];
filterButtons.forEach(button=>button.addEventListener('click',()=>{
  const filter=button.dataset.filter;let count=0;
  filterButtons.forEach(b=>b.setAttribute('aria-pressed',String(b===button)));
  expertiseCards.forEach(card=>{const show=filter==='all'||card.dataset.category.split(' ').includes(filter);card.hidden=!show;if(show)count++;});
  document.querySelector('.expertise-grid').classList.toggle('is-filtered',filter!=='all');
  document.getElementById('expertise-count').textContent=`Showing ${count} ${count===1?'discipline':'disciplines'}${filter==='all'?'.':` for ${button.textContent.trim()}.`}`;
  updateProgress();
}));
document.querySelectorAll('details').forEach(detail=>detail.addEventListener('toggle',updateProgress));
if('IntersectionObserver' in window){
  const activeObserver=new IntersectionObserver(entries=>entries.forEach(entry=>{if(entry.isIntersecting){navigation.querySelectorAll('a').forEach(a=>{const active=a.hash==='#'+entry.target.id;a.classList.toggle('active',active);if(active)a.setAttribute('aria-current','location');else a.removeAttribute('aria-current');});}}),{rootMargin:'-20% 0px -60% 0px'});
  document.querySelectorAll('#work,#experience,#expertise,#contact').forEach(section=>activeObserver.observe(section));
  if(!matchMedia('(prefers-reduced-motion: reduce)').matches){
    const reveals=new IntersectionObserver(entries=>entries.forEach(entry=>{if(entry.isIntersecting){entry.target.classList.add('in-view');reveals.unobserve(entry.target);}}),{threshold:.06});
    document.querySelectorAll('.section-heading,.featured-project,.work-card,.career-intro,.credentials-header').forEach(element=>{element.classList.add('will-reveal');reveals.observe(element);});
  }
}
