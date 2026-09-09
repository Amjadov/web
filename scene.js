/* Locally served Three.js r128. No remote models, textures, or rendering services. */
(() => {
  'use strict';
  const host=document.getElementById('scene-host');
  const sculpture=document.getElementById('sculpture');
  const motionButton=document.querySelector('.motion-toggle');
  const modeButtons=[...document.querySelectorAll('[data-scene]')];
  const motionPreference=matchMedia('(prefers-reduced-motion: reduce)');
  let renderer,frame=0,visible=true,paused=motionPreference.matches,lost=false;
  function fallback(){cancelAnimationFrame(frame);sculpture.classList.remove('scene-ready');host.removeAttribute('tabindex');host.hidden=true;document.querySelector('.scene-toolbar').hidden=true;document.getElementById('scene-instruction').textContent='';document.getElementById('scene-title').textContent='The architecture of impact';}
  try {
    if(!window.THREE){fallback();return;}
    const T=window.THREE;
    renderer=new T.WebGLRenderer({alpha:true,antialias:true,powerPreference:'low-power'});
    renderer.setPixelRatio(Math.min(devicePixelRatio,1.65));
    renderer.outputEncoding=T.sRGBEncoding;
    renderer.toneMapping=T.ACESFilmicToneMapping;
    renderer.toneMappingExposure=1.25;
    host.appendChild(renderer.domElement);
    renderer.domElement.setAttribute('aria-hidden','true');
    const scene=new T.Scene();
    const camera=new T.PerspectiveCamera(39,1,.1,80);
    camera.position.set(0,.2,10.2);
    camera.lookAt(0,0,0);
    // A studio reflection environment gives the sculpture real material depth.
    const studio=document.createElement('canvas');studio.width=2048;studio.height=1024;
    const ctx=studio.getContext('2d');ctx.fillStyle='#242626';ctx.fillRect(0,0,2048,1024);
    const top=ctx.createLinearGradient(0,0,0,1024);top.addColorStop(0,'#777c7a');top.addColorStop(.4,'#222323');top.addColorStop(1,'#080909');ctx.fillStyle=top;ctx.fillRect(0,0,2048,1024);
    ctx.fillStyle='#ffffff';ctx.fillRect(150,80,210,760);ctx.fillRect(1060,50,520,210);
    ctx.fillStyle='#d0e0e1';ctx.fillRect(1740,130,70,640);
    ctx.fillStyle='#ff8954';ctx.fillRect(650,300,250,600);
    const environment=new T.CanvasTexture(studio);environment.mapping=T.EquirectangularReflectionMapping;environment.encoding=T.sRGBEncoding;
    const pmrem=new T.PMREMGenerator(renderer);const envTarget=pmrem.fromEquirectangular(environment);scene.environment=envTarget.texture;environment.dispose();pmrem.dispose();
    scene.add(new T.HemisphereLight(0xffffff,0x292522,.7));
    const key=new T.DirectionalLight(0xffefe1,2.3);key.position.set(-3,5,5);scene.add(key);
    const rim=new T.DirectionalLight(0xff6d35,2);rim.position.set(4,-1,-2);scene.add(rim);
    const fill=new T.DirectionalLight(0xc7e4f0,.9);fill.position.set(2,2,4);scene.add(fill);
    const silver=new T.MeshPhysicalMaterial({color:0xcacac5,metalness:1,roughness:.21,clearcoat:1,clearcoatRoughness:.16,envMapIntensity:1.5});
    const copper=new T.MeshPhysicalMaterial({color:0xff8050,metalness:.83,roughness:.22,clearcoat:1,envMapIntensity:1.4});
    const dark=new T.MeshPhysicalMaterial({color:0x434746,metalness:.9,roughness:.24,clearcoat:1});
    const system=new T.Group();scene.add(system);
    const connected=new T.Group();
    const knot=new T.Mesh(new T.TorusKnotGeometry(1.5,.44,220,32,2,3),silver);knot.rotation.set(.25,-.25,-.3);connected.add(knot);
    const seam=new T.Mesh(new T.TorusKnotGeometry(1.5,.012,220,8,2,3),copper);seam.rotation.copy(knot.rotation);seam.scale.setScalar(1.21);connected.add(seam);
    const core=new T.Mesh(new T.IcosahedronGeometry(.38,3),copper);connected.add(core);
    const architecture=new T.Group();
    const archParts=[];
    for(let x=-1;x<=1;x++)for(let y=-1;y<=1;y++)for(let z=-1;z<=1;z++){
      if(x===0&&y===0&&z===0)continue;
      const mesh=new T.Mesh(new T.BoxGeometry(.68,.68,.68),((x+y+z)%3===0)?copper:((x+y+z)%2===0?silver:dark));
      mesh.position.set(x*1.12,y*1.12,z*1.12);mesh.userData.origin=mesh.position.clone();architecture.add(mesh);archParts.push(mesh);
    }
    architecture.rotation.set(.3,.5,.15);
    const architectureCore=new T.Mesh(new T.IcosahedronGeometry(.42,1),copper);architecture.add(architectureCore);
    const delivery=new T.Group();
    const deliveryRings=[];
    for(let i=0;i<7;i++){
      const ring=new T.Mesh(new T.TorusGeometry(1.45+i*.105,.09,16,120),i===3?copper:silver);
      ring.rotation.set(Math.PI*.43,i*.22,0);ring.position.y=(i-3)*.28;delivery.add(ring);deliveryRings.push(ring);
    }
    delivery.add(new T.Mesh(new T.IcosahedronGeometry(.43,2),copper));
    const forms={connect:connected,architect:architecture,deliver:delivery};
    for(const [name,group] of Object.entries(forms)){group.scale.setScalar(name==='connect'?1:.001);group.visible=name==='connect';system.add(group);}
    const orbit=new T.Mesh(new T.TorusGeometry(3.05,.006,6,180),new T.MeshBasicMaterial({color:0x8e6c52,transparent:true,opacity:.4}));
    orbit.rotation.set(1.17,-.35,-.2);scene.add(orbit);
    const orbitDot=new T.Mesh(new T.SphereGeometry(.045,12,12),new T.MeshBasicMaterial({color:0xff956e}));scene.add(orbitDot);
    const particleGeometry=new T.BufferGeometry();const particles=[];
    // Deterministic positions keep this composition stable across loads.
    for(let i=0;i<64;i++){const a=i*2.399963;const r=2.6+(Math.sin(i*91.1)*.5+.5)*1.2;particles.push(Math.cos(a)*r,Math.sin(a)*r*.8,Math.sin(i*12.6)*1.5-1);}
    particleGeometry.setAttribute('position',new T.Float32BufferAttribute(particles,3));
    scene.add(new T.Points(particleGeometry,new T.PointsMaterial({color:0xaaa08b,size:.018,transparent:true,opacity:.45})));
    let selected='connect',rotX=.08,rotY=-.25,dragging=false,lastX=0,lastY=0,tick=0,lastTime=0;
    function updateMotionLabel(){motionButton.setAttribute('aria-pressed',String(paused));motionButton.setAttribute('aria-label',paused?'Play sculpture animation':'Pause sculpture animation');motionButton.title=paused?'Play animation':'Pause animation';motionButton.textContent=paused?'▶':'Ⅱ';}
    function render(now=0){
      frame=0;if(lost||!visible||document.hidden)return;
      const delta=Math.min((now-lastTime)/1000,.035)||0;lastTime=now;
      if(!paused&&!dragging){tick+=delta;rotY+=delta*.115;}
      system.rotation.x+=(rotX-system.rotation.x)*.12;system.rotation.y+=(rotY-system.rotation.y)*.12;
      system.position.y=paused?0:Math.sin(tick*.6)*.07;
      let transitioning=false;
      for(const [name,group] of Object.entries(forms)){
        const target=name===selected?1:.001;
        const next=motionPreference.matches?target:group.scale.x+(target-group.scale.x)*.13;
        group.scale.setScalar(Math.abs(next-target)<.001?target:next);group.visible=group.scale.x>.008;
        if(Math.abs(group.scale.x-target)>.001)transitioning=true;
      }
      if(!paused){archParts.forEach((p,i)=>{p.position.copy(p.userData.origin).multiplyScalar(1+Math.sin(tick*.65+i*.12)*.025);});deliveryRings.forEach((r,i)=>r.rotation.z=Math.sin(tick*.5+i*.4)*.18);}
      orbitDot.position.set(Math.cos(tick*.24)*3.05,Math.sin(tick*.24)*3.05,0).applyEuler(orbit.rotation);
      renderer.render(scene,camera);
      sculpture.classList.add('scene-ready');
      if(!paused||transitioning||Math.abs(system.rotation.x-rotX)>.001||Math.abs(system.rotation.y-rotY)>.001)frame=requestAnimationFrame(render);
    }
    function requestRender(){if(!frame&&visible&&!document.hidden&&!lost)frame=requestAnimationFrame(render);}
    function resize(){const r=host.getBoundingClientRect();if(r.width<=0||r.height<=0)return;renderer.setSize(r.width,r.height,false);camera.aspect=r.width/r.height;camera.position.z=camera.aspect<1?11.4:10.2;camera.updateProjectionMatrix();requestRender();}
    modeButtons.forEach(b=>b.addEventListener('click',()=>{selected=b.dataset.scene;modeButtons.forEach(m=>m.setAttribute('aria-pressed',String(m===b)));document.getElementById('scene-title').textContent={connect:'The connected whole',architect:'Structure from complexity',deliver:'Momentum with purpose'}[selected];requestRender();}));
    motionButton.addEventListener('click',()=>{paused=!paused;updateMotionLabel();requestRender();});
    motionPreference.addEventListener('change',e=>{paused=e.matches;updateMotionLabel();requestRender();});
    host.addEventListener('pointerdown',e=>{if(e.button!==0)return;dragging=true;lastX=e.clientX;lastY=e.clientY;host.setPointerCapture(e.pointerId);});
    host.addEventListener('pointermove',e=>{if(!dragging)return;rotY+=(e.clientX-lastX)*.007;rotX=Math.max(-1.15,Math.min(1.15,rotX+(e.clientY-lastY)*.005));lastX=e.clientX;lastY=e.clientY;requestRender();});
    const endDrag=()=>{dragging=false;requestRender();};host.addEventListener('pointerup',endDrag);host.addEventListener('pointercancel',endDrag);host.addEventListener('lostpointercapture',endDrag);
    host.addEventListener('keydown',e=>{if(!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Home'].includes(e.key))return;e.preventDefault();if(e.key==='ArrowLeft')rotY-=.2;if(e.key==='ArrowRight')rotY+=.2;if(e.key==='ArrowUp')rotX=Math.max(-1.15,rotX-.2);if(e.key==='ArrowDown')rotX=Math.min(1.15,rotX+.2);if(e.key==='Home'){rotX=.08;rotY=-.25;}requestRender();});
    new ResizeObserver(resize).observe(host);
    new IntersectionObserver(entries=>{visible=entries[0].isIntersecting;if(visible){lastTime=performance.now();requestRender();}else{cancelAnimationFrame(frame);frame=0;}},{threshold:0}).observe(host);
    document.addEventListener('visibilitychange',()=>{if(document.hidden){cancelAnimationFrame(frame);frame=0;}else{lastTime=performance.now();requestRender();}});
    renderer.domElement.addEventListener('webglcontextlost',e=>{e.preventDefault();lost=true;fallback();});
    updateMotionLabel();resize();
  }catch(error){console.warn('3D view unavailable; displaying portfolio artwork.',error);if(renderer)renderer.dispose();fallback();}
})();
