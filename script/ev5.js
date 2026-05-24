// npm install docx
// node generate_exam_doc_even.js

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, LevelFormat, PageBreak, BorderStyle
  } = require('docx');
  const fs = require('fs');
  
  // ─── helpers ────────────────────────────────────────────────────────────────
  
  function h1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text, bold: true, size: 32, font: 'Arial' })],
      spacing: { before: 400, after: 200 },
    });
  }
  
  function h2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text, bold: true, size: 28, font: 'Arial' })],
      spacing: { before: 300, after: 160 },
    });
  }
  
  function para(text) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 100, after: 100, line: 276 },
      children: [new TextRun({ text, font: 'Arial', size: 24 })],
    });
  }
  
  function bullet(text) {
    return new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, font: 'Arial', size: 24 })],
    });
  }
  
  function code(text) {
    return new Paragraph({
      spacing: { before: 60, after: 60 },
      indent: { left: 720 },
      children: [new TextRun({ text, font: 'Courier New', size: 20 })],
    });
  }
  
  function pageBreak() {
    return new Paragraph({ children: [new PageBreak()] });
  }
  
  function divider() {
    return new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 1 } },
      spacing: { before: 200, after: 200 },
      children: [],
    });
  }
  
  // ─── content ─────────────────────────────────────────────────────────────────
  
  const sections = [];
  
  // ── Q2 ───────────────────────────────────────────────────────────────────────
  sections.push(
    h1('Питання 2. Класифікація паралельних обчислювальних систем'),
    divider(),
  
    h2('1. Загальний огляд'),
    para('Паралельні обчислювальні системи класифікують за різними критеріями: архітектурою пам\'яті, топологією з\'єднань, моделлю програмування та типом паралелізму. Розуміння цієї класифікації є фундаментом для вибору правильного підходу до розпаралелення задачі.'),
  
    h2('2. Таксономія Фліна (Flynn, 1966)'),
    para('Найвідоміша класифікація за потоками команд (Instruction Streams, IS) та потоками даних (Data Streams, DS):'),
    bullet('SISD (Single IS, Single DS) — класичний послідовний процесор фон Неймана. Одна команда — один операнд. Приклад: однопроцесорний ПК без векторних розширень.'),
    bullet('SIMD (Single IS, Multiple DS) — одна команда над вектором даних одночасно. Приклади: AVX/SSE у сучасних CPU, GPU-ядра (тисячі CUDA-ядер виконують одну функцію над різними пікселями), DSP-процесори.'),
    bullet('MISD (Multiple IS, Single DS) — кілька команд над одними даними. Теоретична категорія; застосовується у відмовостійких системах (fault-tolerant), де кілька процесорів незалежно перевіряють той самий результат.'),
    bullet('MIMD (Multiple IS, Multiple DS) — найпоширеніший клас паралельних систем. Кожен процесор виконує свою програму над своїми даними. Охоплює багатоядерні CPU, кластери, суперкомп\'ютери.'),
  
    h2('3. Класифікація MIMD за пам\'яттю'),
    para('Shared Memory (SMP/NUMA): всі процесори мають доступ до спільного адресного простору. Комунікація через запис/читання змінних. Синхронізація через locks, semaphores. Приклад: багатоядерний Intel Xeon. У Java це основна модель — потоки поділяють heap.'),
    para('Distributed Memory (MPP): кожен вузол має власну пам\'ять. Комунікація лише через явний обмін повідомленнями (MPI). Приклад: кластер з InfiniBand, суперкомп\'ютер Frontier.'),
    para('Hybrid (кластери з багатоядерними вузлами): всередині вузла — shared memory (OpenMP або Java threads), між вузлами — message passing (MPI). Більшість сучасних HPC-систем саме такого типу.'),
  
    h2('4. Класифікація за зв\'язністю процесорів'),
    bullet('Тісно зв\'язані (Tightly Coupled) — спільна пам\'ять, швидка взаємодія, наприклад через шину або кросбар'),
    bullet('Слабко зв\'язані (Loosely Coupled) — кожен вузол автономний, зв\'язок через мережу (LAN, InfiniBand)'),
  
    h2('5. Класифікація за топологією міжз\'єднань'),
    bullet('Bus (шина) — проста, дешева, але обмежена пропускна здатність'),
    bullet('Ring (кільце) — кожен вузол з\'єднаний з двома сусідами'),
    bullet('Mesh / Torus (2D/3D решітка) — Blue Gene, Cray XE6'),
    bullet('Hypercube — 2^N вузлів, діаметр N; висока зв\'язність'),
    bullet('Fat-Tree — дерево з потовщеними гілками вгорі; стандарт InfiniBand-кластерів'),
    bullet('Butterfly / Omega / Beneš — перемикальні мережі для паралельних комп\'ютерів'),
  
    h2('6. Класифікація за моделлю програмування'),
    bullet('Спільна пам\'ять: Pthreads, Java Threads, OpenMP, Intel TBB'),
    bullet('Передача повідомлень: MPI, PVM'),
    bullet('Гібридна: MPI + OpenMP, MPI + Java ForkJoin'),
    bullet('Потоки даних (dataflow): CUDA, OpenCL, SYCL'),
    bullet('Функціональний паралелізм: Erlang, Haskell par/seq, Scala Akka'),
  
    h2('7. Зв\'язок з Java'),
    para('У Java основна модель — shared-memory MIMD: потоки Thread/Runnable поділяють heap JVM. Для distributed-memory використовують MPI через MPJ Express або Spring Boot REST/gRPC між мікросервісами. ForkJoinPool реалізує MIMD на рівні завдань усередині JVM.'),
    pageBreak(),
  
    // ── Q4 ───────────────────────────────────────────────────────────────────
    h1('Питання 4. Системи із загальною та розподіленою пам\'яттю'),
    divider(),
  
    h2('1. Системи із загальною пам\'яттю (Shared Memory)'),
    para('У цій архітектурі всі процесори (ядра) мають прямий доступ до єдиного фізичного адресного простору оперативної пам\'яті. Взаємодія між потоками — через читання/запис спільних змінних. Не потрібен явний обмін повідомленнями.'),
    para('Різновиди:'),
    bullet('UMA (Uniform Memory Access) — затримка доступу до RAM однакова для всіх ядер. Класичний SMP (Symmetric Multiprocessing). Масштабується погано понад ~16 ядер через конкуренцію на шині.'),
    bullet('NUMA (Non-Uniform Memory Access) — кожен процесор має "локальну" RAM; доступ до "чужої" RAM повільніший через міжз\'єднання QPI/UPI. Сучасні сервери AMD EPYC, Intel Xeon — NUMA. Програміст повинен враховувати NUMA-топологію для максимальної швидкодії.'),
    bullet('CC-NUMA (Cache-Coherent NUMA) — апаратна гарантія когерентності кешів між ядрами. Стандарт для сучасних багатоядерних серверів.'),
    para('Переваги shared memory: простота програмування (Java threads, OpenMP), низька латентність комунікації (запис у RAM замість мережевого пакету), природна підтримка в JVM.'),
    para('Недоліки: важко масштабувати понад кілька десятків ядер (конфлікти кешу, bus contention); один збій руйнує всю систему (немає fault isolation).'),
  
    h2('2. Когерентність кешів'),
    para('Кожне ядро має власний кеш L1/L2. Якщо ядро A записало нове значення в кеш, ядро B не побачить оновлення до flush — проблема когерентності. Протоколи MESI, MESIF, MOESI забезпечують автоматичну когерентність апаратно. У Java модель пам\'яті (JMM) гарантує видимість змін через synchronized або volatile.'),
  
    h2('3. Системи з розподіленою пам\'яттю (Distributed Memory)'),
    para('Кожен вузол (node) має власний процесор і власну пам\'ять. Жоден вузол не може безпосередньо читати/писати пам\'ять іншого. Взаємодія — лише через явний обмін повідомленнями по мережі (MPI_Send/MPI_Recv, gRPC, REST).'),
    para('Переваги: масштабованість до тисяч вузлів (суперкомп\'ютери, хмара); fault isolation (збій одного вузла не руйнує інші); відсутні конфлікти кешів між вузлами.'),
    para('Недоліки: висока латентність комунікації (мкс–мс замість нс); програміст явно управляє розподілом даних; складніше налагоджування.'),
  
    h2('4. Гібридна архітектура'),
    para('Сучасні HPC-кластери — гібрид: всередині кожного вузла shared memory (багатоядерний CPU + GPU), між вузлами — distributed memory (InfiniBand, Ethernet). Програмна модель: OpenMP або Java Threads усередині вузла + MPI між вузлами.'),
  
    h2('5. Порівняння у контексті Java'),
    bullet('Java Threads / ForkJoinPool → shared memory на багатоядерному CPU'),
    bullet('Spring Boot мікросервіси → розподілена пам\'ять (кожен сервіс ізольований, REST/gRPC для взаємодії)'),
    bullet('MPJ Express (MPI для Java) → distributed memory на кластері'),
    bullet('Matrix multiply: shared memory — стрічковий алгоритм з потоками; distributed memory — MPI Scatter/Gather'),
    pageBreak(),
  
    // ── Q6 ───────────────────────────────────────────────────────────────────
    h1('Питання 6. Багатопоточна технологія Java'),
    divider(),
  
    h2('1. Архітектура потоків JVM'),
    para('JVM запускається як один процес ОС. Всередині — кілька потоків (threads), що поділяють heap і метадані класів, але мають власний стек викликів, лічильник команд та регістри. JVM відображає Java-потоки на нативні потоки ОС (1:1 у HotSpot). Планування — відповідальність ОС; JVM лише підказує пріоритетом.'),
  
    h2('2. Три способи створення потоку'),
    para('Спосіб 1 — наслідування Thread:'),
    code('class Worker extends Thread {'),
    code('    @Override public void run() { compute(); }'),
    code('}'),
    code('new Worker().start();'),
    para('Спосіб 2 — реалізація Runnable (рекомендований):'),
    code('Runnable task = () -> compute();'),
    code('new Thread(task, "worker-1").start();'),
    para('Спосіб 3 — Callable + Future (з результатом):'),
    code('ExecutorService es = Executors.newFixedThreadPool(4);'),
    code('Future<Double> f = es.submit(() -> heavyCalc());'),
    code('double result = f.get(); // блокує до готовності'),
  
    h2('3. Пріоритети потоків'),
    para('Встановлюються через setPriority(int): MIN_PRIORITY=1, NORM_PRIORITY=5, MAX_PRIORITY=10. Це лише підказка планувальнику; гарантій немає. На Linux JVM пріоритети відображаються на nice-рівні ОС.'),
  
    h2('4. Синхронізація'),
    para('Keyword synchronized забезпечує взаємне виключення через монітор об\'єкта:'),
    code('synchronized (sharedObject) { /* критична секція */ }'),
    code('public synchronized void increment() { count++; }'),
    para('JMM гарантує happens-before: записи перед виходом зі synchronized видимі потокам, що входять у той самий монітор. volatile гарантує видимість без взаємного виключення.'),
  
    h2('5. java.util.concurrent — високий рівень'),
    bullet('ExecutorService / ThreadPoolExecutor — пул потоків, управління задачами'),
    bullet('ForkJoinPool — work-stealing, divide-and-conquer'),
    bullet('CompletableFuture — асинхронні обчислення з ланцюжками (thenApply, thenCompose)'),
    bullet('ReentrantLock, ReadWriteLock — гнучкіші замки'),
    bullet('Semaphore, CountDownLatch, CyclicBarrier — координація'),
    bullet('ConcurrentHashMap, LinkedBlockingQueue — потокобезпечні колекції'),
    bullet('AtomicInteger, LongAdder — lock-free операції через CAS'),
  
    h2('6. Модель пам\'яті Java (JMM)'),
    para('JMM визначає, коли зміна однієї нитки стає видимою іншій. Без синхронізації — немає гарантій (CPU може кешувати значення, компілятор може перетасовувати операції). Засоби синхронізації: synchronized, volatile, final (поля, записані в конструкторі, гарантовано видимі після завершення конструктора), java.util.concurrent.'),
  
    h2('7. Virtual Threads (Java 21+)'),
    para('Project Loom вводить віртуальні потоки (Virtual Threads) — надлегкі (кілька KB замість ~1MB стеку), керовані JVM, мультиплексовані на нативні потоки. Дозволяють мільйони одночасних потоків. Ідеальні для I/O-bound задач (REST-сервери). Створення: Thread.ofVirtual().start(task).'),
    pageBreak(),
  
    // ── Q8 ───────────────────────────────────────────────────────────────────
    h1('Питання 8. Стани потоку та переходи між станами'),
    divider(),
  
    h2('1. Перелік станів (Thread.State)'),
    bullet('NEW — потік створено (new Thread()), але start() ще не викликано. Не виконується.'),
    bullet('RUNNABLE — потік виконується або готовий до виконання (очікує квант CPU від планувальника). Включає стан "виконується на ядрі" та "стоїть у черзі готових".'),
    bullet('BLOCKED — потік намагається увійти до synchronized-блоку або методу, але монітор зайнятий іншим потоком. Потік не споживає CPU; він стоїть у entry set монітора.'),
    bullet('WAITING — потік чекає безстроково: викликав wait(), join() без таймауту або LockSupport.park(). Прокинеться лише після явного notify()/notifyAll() або unpark().'),
    bullet('TIMED_WAITING — чекає обмежений час: Thread.sleep(ms), wait(ms), join(ms), LockSupport.parkNanos(). Автоматично прокидається після таймауту.'),
    bullet('TERMINATED — метод run() завершився (нормально або через виняток). Стан незворотній; повторний start() кине IllegalThreadStateException.'),
  
    h2('2. Діаграма переходів'),
    para('NEW → RUNNABLE: виклик start().'),
    para('RUNNABLE → BLOCKED: спроба увійти в зайнятий synchronized-блок.'),
    para('BLOCKED → RUNNABLE: власник монітора звільнив його (вийшов зі synchronized).'),
    para('RUNNABLE → WAITING: виклик wait(), join(), LockSupport.park().'),
    para('WAITING → RUNNABLE: виклик notify()/notifyAll() іншим потоком, або join()-потік завершився.'),
    para('RUNNABLE → TIMED_WAITING: Thread.sleep(n), wait(n), join(n).'),
    para('TIMED_WAITING → RUNNABLE: таймаут минув або умова виконалась достроково.'),
    para('RUNNABLE → TERMINATED: повернення з run() або необроблений виняток.'),
  
    h2('3. Приклад перевірки стану'),
    code('Thread t = new Thread(() -> {'),
    code('    try { Thread.sleep(5000); }'),
    code('    catch (InterruptedException e) { Thread.currentThread().interrupt(); }'),
    code('});'),
    code('System.out.println(t.getState()); // NEW'),
    code('t.start();'),
    code('Thread.sleep(100);'),
    code('System.out.println(t.getState()); // TIMED_WAITING'),
    code('t.join();'),
    code('System.out.println(t.getState()); // TERMINATED'),
  
    h2('4. Важливі тонкощі'),
    para('wait() звільняє монітор (потік переходить у wait set), тоді як sleep() монітор НЕ звільняє. Після notify() потік переходить з wait set до entry set — він ще не RUNNABLE, а BLOCKED, поки не отримає монітор знову. yield() залишає потік у RUNNABLE, але підказує планувальнику дати можливість іншим.'),
  
    h2('5. Моніторинг стану потоків'),
    para('Для аналізу виробничих систем: Thread.getAllStackTraces() повертає стан і трасування стеку всіх потоків. jstack <pid> — утиліта JDK для дампу потоків. VisualVM та Java Flight Recorder (JFR) — графічний аналіз. Стан BLOCKED кількох потоків на одному моніторі — сигнал можливого deadlock або высокого lock contention.'),
    pageBreak(),
  
    // ── Q10 ──────────────────────────────────────────────────────────────────
    h1('Питання 10. Алгоритми паралельного множення матриць'),
    divider(),
  
    h2('1. Постановка задачі'),
    para('Множення матриць A (N×M) та B (M×K) → C (N×K), де C[i][j] = Σ A[i][t]*B[t][j]. Послідовна складність O(N³). Для великих матриць (N=1000+) розпаралелення дає суттєвий виграш. Алгоритми розрізняються за: способом декомпозиції даних, кеш-ефективністю та пристосованістю до shared/distributed memory.'),
  
    h2('2. Наївний рядковий алгоритм'),
    para('Кожен потік t обчислює рядки [start_t, end_t) матриці C. Проста реалізація, добра для shared memory (Java Threads). Кеш-неефективна при стовпчиковому читанні B (cache miss).'),
    code('for (int i = startRow; i < endRow; i++)'),
    code('  for (int j = 0; j < K; j++) {'),
    code('    double s = 0;'),
    code('    for (int k = 0; k < M; k++) s += A[i][k] * B[k][j];'),
    code('    C[i][j] = s;'),
    code('  }'),
  
    h2('3. Стрічковий алгоритм з транспозицією B'),
    para('Транспонуємо B → BT один раз. Тоді C[i][j] = Σ A[i][k]*BT[j][k] — обидва масиви читаються рядково (cache-friendly). Ефективний для shared memory, прискорення може бути в 2–4 рази порівняно з наївним за рахунок кешу.'),
  
    h2('4. Блоковий алгоритм (tile/block)'),
    para('Матриці розбиваються на блоки розміру b×b. Обчислюємо C[I][J] += A[I][K] * B[K][J] блоками. Весь блок поміщається у L1/L2 кеш → драматичне зменшення cache miss. Розмір блоку b обирається під розмір кешу (~32–128 елементів).'),
    code('for (int I = 0; I < N; I += b)'),
    code('  for (int J = 0; J < K; J += b)'),
    code('    for (int KK = 0; KK < M; KK += b)'),
    code('      multiplyBlock(A, B, C, I, J, KK, b);'),
  
    h2('5. Алгоритм Штрасена'),
    para('Рекурсивний алгоритм: замість 8 множень 2×2 блоків — 7 (через хитрі лінійні комбінації). Складність O(N^2.807). Для великих матриць (N>512) перевершує наївний. Складний у реалізації, є чисельна нестабільність.'),
  
    h2('6. Алгоритм Фокса (для distributed memory)'),
    para('Процеси розташовані у решітці √p × √p. На кожному кроці 1..√p: broadcast блоку A по рядку → локальне множення → shift блоку B вгору по стовпцю. √p кроків, кожен — broadcast + зсув + множення блоків. Ефективний для MPI-кластерів.'),
  
    h2('7. Алгоритм Кеннона'),
    para('Аналог Фокса але з зсувами замість broadcast (менший трафік). Початкове вирівнювання A і B, потім p кроків зсуву + множення. Мінімальні комунікаційні витрати O(N²/√p).'),
  
    h2('8. Spring Boot для множення матриць'),
    para('REST-контролер приймає матриці, розподіляє рядки по потоках (ForkJoin або FixedThreadPool), агрегує результат і повертає. Клієнт отримує відповідь асинхронно через CompletableFuture або синхронно через @PostMapping.'),
    pageBreak(),
  
    // ── Q12 ──────────────────────────────────────────────────────────────────
    h1('Питання 12. Алгоритм Фокса паралельного множення матриць'),
    divider(),
  
    h2('1. Передумови'),
    para('Алгоритм Фокса (Fox algorithm) — блоковий паралельний алгоритм для distributed-memory систем (MPI). Розроблений для решітки процесів p = q² (q = √p). Ефективно використовує колективні MPI-операції.'),
  
    h2('2. Підготовка: розбиття на блоки'),
    para('Матриці A і B (розмір N×N) розбиваються на блоки розміру (N/q)×(N/q). Блок A[i][j] та B[i][j] спочатку знаходяться у процесі (i,j) решітки. Результат C[i][j] накопичується у процесі (i,j).'),
  
    h2('3. Кроки алгоритму (q ітерацій, k = 0..q-1)'),
    para('На кроці k:'),
    bullet('Крок 3.1: Broadcast блоку A[i][(i+k) mod q] по всьому рядку i решітки. Кожен процес рядка i отримує один і той самий блок A.'),
    bullet('Крок 3.2: Локальне множення: C[i][j] += A_broadcast * B_local[i][j].'),
    bullet('Крок 3.3: Cyclic shift блоків B вгору по кожному стовпцю решітки: B[i][j] переходить до процесу ((i-1+q) mod q, j).'),
    para('Після q кроків кожен процес (i,j) має повний підсумок C[i][j] = Σ A[i][k]*B[k][j].'),
  
    h2('4. MPI-реалізація (схематично)'),
    code('// Визначення координат у решітці:'),
    code('MPI_Cart_create(comm, 2, dims, periods, 1, &cartComm);'),
    code('MPI_Cart_coords(cartComm, rank, 2, coords); // coords = {row, col}'),
    code('MPI_Comm_split(cartComm, coords[0], coords[1], &rowComm);'),
    code('MPI_Comm_split(cartComm, coords[1], coords[0], &colComm);'),
    code(''),
    code('for (int k = 0; k < q; k++) {'),
    code('  // Визначити root для broadcast у рядку:'),
    code('  int root = (coords[0] + k) % q;'),
    code('  if (coords[1] == root) memcpy(tempA, localA, blockSize);'),
    code('  MPI_Bcast(tempA, blockSize, MPI_DOUBLE, root, rowComm);'),
    code('  // Локальне множення:'),
    code('  matMul(tempA, localB, localC, n/q);'),
    code('  // Cyclic shift B вгору:'),
    code('  int dest   = (coords[0] - 1 + q) % q;'),
    code('  int source = (coords[0] + 1) % q;'),
    code('  MPI_Sendrecv_replace(localB, blockSize, MPI_DOUBLE,'),
    code('                       dest, 0, source, 0, colComm, &status);'),
    code('}'),
  
    h2('5. Аналіз складності'),
    bullet('Обчислення: O(N³/p) на процес — ідеальне балансування'),
    bullet('Комунікації: q кроків × (broadcast по рядку + shift по стовпцю)'),
    bullet('Broadcast по рядку з q процесів: O(log q * (N/q)²)'),
    bullet('Shift по стовпцю: O((N/q)²) — peer-to-peer'),
    bullet('Загально: T_comm = O(N² log(√p) / p)'),
  
    h2('6. Порівняння з іншими алгоритмами'),
    bullet('Фокс vs наївний MPI scatter/gather: Фокс — менший обсяг broadcast, краще масштабується'),
    bullet('Фокс vs Кеннон: Фокс простіший, але більший трафік (broadcast замість shift)'),
    bullet('Фокс vs стрічковий (shared memory): стрічковий простіший для Java threads, Фокс — для MPI-кластерів'),
    pageBreak(),
  
    // ── Q14 ──────────────────────────────────────────────────────────────────
    h1('Питання 14. Призупинка та відновлення роботи потоку. Переривання дії потоку'),
    divider(),
  
    h2('1. Thread.sleep() — тимчасова призупинка'),
    para('Thread.sleep(ms) переводить поточний потік у стан TIMED_WAITING на вказану кількість мілісекунд. Потік не споживає CPU. Важливо: sleep() НЕ звільняє утримувані монітори (на відміну від wait()). Завжди може бути перерваний через InterruptedException:'),
    code('try {'),
    code('    Thread.sleep(2000); // пауза 2 секунди'),
    code('} catch (InterruptedException e) {'),
    code('    Thread.currentThread().interrupt(); // відновити прапор!'),
    code('    return; // коректне завершення'),
    code('}'),
  
    h2('2. wait() / notify() — умовне очікування'),
    para('Викликаються тільки всередині synchronized-блоку. wait() звільняє монітор і переводить потік у WAITING. notify() прокидає один із потоків у wait set. notifyAll() прокидає всіх:'),
    code('synchronized (lock) {'),
    code('    while (!condition) lock.wait(); // re-check після пробудження!'),
    code('    doWork();'),
    code('    lock.notifyAll();'),
    code('}'),
    para('Чому while, а не if? Spurious wakeup — потік може прокинутись без notify. Завжди перевіряти умову в циклі.'),
  
    h2('3. join() — очікування іншого потоку'),
    para('t.join() блокує поточний потік (WAITING) до завершення потоку t. t.join(ms) — TIMED_WAITING, прокидається після ms незалежно від завершення t:'),
    code('Thread[] workers = new Thread[4];'),
    code('for (int i = 0; i < 4; i++) { workers[i] = new Thread(task); workers[i].start(); }'),
    code('for (Thread w : workers) w.join(); // чекати всіх'),
  
    h2('4. LockSupport.park() / unpark()'),
    para('Нижчий рівень ніж wait/notify. park() призупиняє потік без монітора. unpark(thread) дозволяє конкретному потоку продовжити. Використовується всередині ReentrantLock та інших java.util.concurrent примітивів.'),
  
    h2('5. Механізм переривання (Interruption)'),
    para('interrupt() встановлює прапор переривання у потоці. Сам по собі нікуди не переміщує потік — він лише сигналізує. Потік сам відповідає за перевірку і реакцію. Методи, що кидають InterruptedException (sleep, wait, join), реагують негайно і скидають прапор — тому треба відновити його через Thread.currentThread().interrupt() у catch-блоці:'),
    code('// Коректна структура cancellable-потоку:'),
    code('public void run() {'),
    code('    while (!Thread.currentThread().isInterrupted()) {'),
    code('        try {'),
    code('            processNextItem();'),
    code('            Thread.sleep(100);'),
    code('        } catch (InterruptedException e) {'),
    code('            Thread.currentThread().interrupt();'),
    code('            break; // виходимо з циклу'),
    code('        }'),
    code('    }'),
    code('    cleanup();'),
    code('}'),
  
    h2('6. Застаріле: suspend() / resume() / stop()'),
    para('Методи Thread.suspend(), resume(), stop() є deprecated з Java 1.2. stop() небезпечний — залишає об\'єкти у неузгодженому стані, звільняючи монітори між рядками synchronized. Правильна альтернатива — volatile boolean або AtomicBoolean прапор + interrupt().'),
    pageBreak(),
  
    // ── Q16 ──────────────────────────────────────────────────────────────────
    h1('Питання 16. Синхронізація в паралельних обчисленнях'),
    divider(),
  
    h2('1. Що таке синхронізація і навіщо вона потрібна'),
    para('Синхронізація — механізми координації паралельних потоків/процесів, що забезпечують: взаємне виключення (лише один потік у критичній секції), упорядкованість операцій (happens-before), видимість змін між потоками.'),
    para('Без синхронізації виникають: race condition (результат залежить від порядку виконання), data race (паралельний запис/читання без синхронізації), visibility problem (потік бачить застарілі значення з кешу).'),
  
    h2('2. Мютекс (Mutual Exclusion Lock)'),
    para('Найбасовіший примітив. Лише один потік тримає мютекс одночасно; інші блокуються. У Java: keyword synchronized (неявний монітор), ReentrantLock (явний):'),
    code('ReentrantLock mutex = new ReentrantLock();'),
    code('mutex.lock();'),
    code('try { criticalSection(); }'),
    code('finally { mutex.unlock(); }'),
  
    h2('3. Семафор'),
    para('Підраховуючий семафор (counting semaphore) дозволяє N потокам одночасно. acquire() зменшує лічильник (або блокує при 0); release() збільшує. Двійковий семафор (N=1) ≡ мютекс. Використання: обмеження числа одночасних з\'єднань до БД:'),
    code('Semaphore dbPool = new Semaphore(10);'),
    code('dbPool.acquire(); // отримати дозвіл'),
    code('try { useDatabase(); } finally { dbPool.release(); }'),
  
    h2('4. Бар\'єрна синхронізація'),
    para('Усі потоки зустрічаються в точці синхронізації; жоден не просувається далі, поки всі не досягли бар\'єру. CyclicBarrier у Java:'),
    code('CyclicBarrier barrier = new CyclicBarrier(numThreads, () -> {'),
    code('    System.out.println("Всі завершили фазу, починаємо наступну");'),
    code('});'),
    code('// Кожен потік після фази:'),
    code('barrier.await();'),
  
    h2('5. Монітор (Java synchronized + wait/notify)'),
    para('Монітор об\'єднує мютекс і умовну змінну. synchronized гарантує взаємне виключення; wait/notify — умовне очікування. Класична реалізація черги (Producer-Consumer):'),
    code('class BoundedQueue<T> {'),
    code('    private final LinkedList<T> queue = new LinkedList<>();'),
    code('    private final int limit;'),
    code('    synchronized void put(T item) throws InterruptedException {'),
    code('        while (queue.size() == limit) wait();'),
    code('        queue.add(item); notifyAll();'),
    code('    }'),
    code('    synchronized T take() throws InterruptedException {'),
    code('        while (queue.isEmpty()) wait();'),
    code('        T item = queue.poll(); notifyAll(); return item;'),
    code('    }'),
    code('}'),
  
    h2('6. Синхронізація в MPI'),
    para('MPI_Barrier(comm) — бар\'єр для всіх процесів комунікатора: жоден не виходить, поки всі не дійшли. Використовується для вимірювання часу (після ініціалізації до MPI_Wtime). Блокуючі MPI_Send/MPI_Recv — неявна синхронізація між двома процесами.'),
  
    h2('7. Синхронізація пам\'яті (Memory Barriers / Fences)'),
    para('На рівні CPU: store barrier (flush записи в RAM), load barrier (інвалідація кешу для свіжого читання). У Java: synchronized та volatile містять неявні бар\'єри пам\'яті. Без них компілятор і CPU можуть переставляти операції (instruction reordering).'),
    pageBreak(),
  
    // ── Q18 ──────────────────────────────────────────────────────────────────
    h1('Питання 18. Блокування об\'єкту'),
    divider(),
  
    h2('1. Монітор об\'єкта Java'),
    para('Кожен об\'єкт у Java має асоційований монітор (implicit lock). Монітор — це exclusive lock: у будь-який момент лише один потік може його тримати. Реалізується у JVM через mark word у заголовку об\'єкта (biased/thin/fat lock — залежно від конкуренції).'),
  
    h2('2. Захоплення монітора через synchronized'),
    para('synchronized method: захоплює монітор this (або Class для статичних). synchronized block: явно вказуємо об\'єкт-монітор. Вхід: якщо монітор вільний — потік захоплює і продовжує; якщо зайнятий — переходить у BLOCKED.'),
    code('// Різні об\'єкти-монітори — різні замки (не заважають один одному):'),
    code('synchronized (lockA) { /* операції над A */ }'),
    code('synchronized (lockB) { /* операції над B */ } // може виконуватись паралельно'),
  
    h2('3. Реентерабельність (Reentrancy)'),
    para('Java monitors і ReentrantLock — реентерабельні: потік, що вже тримає монітор об\'єкта, може входити у інший synchronized-блок того ж об\'єкта без deadlock. Лічильник входів збільшується; монітор звільняється лише коли лічильник == 0:'),
    code('synchronized void outer() {'),
    code('    inner(); // ОК! Той самий потік захоплює монітор повторно'),
    code('}'),
    code('synchronized void inner() { /* ... */ }'),
  
    h2('4. Object.wait() / notify() — умовна змінна на моніторі'),
    para('wait() тимчасово звільняє монітор і вміщує потік у wait set об\'єкта. notify() переміщує один потік з wait set до entry set (але не дає монітор одразу). notifyAll() — всіх. Потік у wait set не споживає CPU.'),
  
    h2('5. intrinsic lock vs ReentrantLock'),
    para('Intrinsic lock (synchronized): простий синтаксис, автоматичне звільнення при виході зі synchronized (у т.ч. при виключенні), але немає tryLock, немає справедливої черги, немає перемикання між умовними змінними.'),
    para('ReentrantLock: tryLock(timeout), lockInterruptibly(), справедливий режим (fairness=true), кілька Condition об\'єктів (newCondition()), але треба явно unlock() у finally.'),
    code('ReentrantLock lock = new ReentrantLock(true); // fair lock'),
    code('Condition notFull = lock.newCondition();'),
    code('Condition notEmpty = lock.newCondition();'),
    code('// Producer:'),
    code('lock.lock(); try {'),
    code('    while (full) notFull.await();'),
    code('    add(); notEmpty.signal();'),
    code('} finally { lock.unlock(); }'),
  
    h2('6. Оптимізації JVM (lock inflation)'),
    bullet('Biased locking: якщо лише один потік використовує об\'єкт — немає CAS, лише mark word'),
    bullet('Thin lock (lightweight): при конкуренції — CAS у mark word'),
    bullet('Fat lock (heavyweight): при тривалій конкуренції — реальний OS mutex через futex'),
    bullet('Lock elimination: JIT видаляє непотрібні synchronized при аналізі escape'),
    bullet('Lock coarsening: JIT об\'єднує кілька суміжних synchronized у один'),
    pageBreak(),
  
    // ── Q20 ──────────────────────────────────────────────────────────────────
    h1('Питання 20. Проблеми управління потоками: дедлок та інші небажані стани'),
    divider(),
  
    h2('1. Deadlock (взаємне блокування)'),
    para('Deadlock — ситуація, коли два або більше потоки нескінченно чекають один одного. Умови Coffman (всі чотири мають виконуватись одночасно): взаємне виключення (ресурс тримається ексклюзивно), утримання і очікування (потік тримає ресурс і чекає інший), без примусового вилучення (ресурс не можна відібрати), кругове очікування (P1 чекає P2, P2 чекає P1).'),
    code('// Класичний deadlock:'),
    code('// Потік 1: lock(A), потім lock(B)'),
    code('// Потік 2: lock(B), потім lock(A) → обидва заблоковані назавжди'),
  
    h2('2. Виявлення deadlock'),
    para('jstack <pid> або VisualVM — JVM вміє детектувати deadlock через аналіз утримуваних і очікуваних моніторів. У логах: "Found one Java-level deadlock". ThreadMXBean.findDeadlockedThreads() програмно.'),
  
    h2('3. Запобігання deadlock'),
    bullet('Впорядкування замків: завжди захоплювати у одному порядку (наприклад, за id об\'єкта)'),
    bullet('tryLock з таймаутом: якщо не вдалось — відпустити все і спробувати пізніше'),
    bullet('Мінімізувати час утримання замку і кількість одночасно утримуваних'),
    bullet('Lock-free структури даних (ConcurrentHashMap, AtomicInteger)'),
    code('// tryLock запобігання:'),
    code('if (lock1.tryLock(500, MILLISECONDS)) {'),
    code('  try {'),
    code('    if (lock2.tryLock(500, MILLISECONDS)) {'),
    code('      try { doWork(); } finally { lock2.unlock(); }'),
    code('    }'),
    code('  } finally { lock1.unlock(); }'),
    code('}'),
  
    h2('4. Livelock'),
    para('Потоки не заблоковані, але постійно реагують один на одного і не рухаються вперед. Аналогія: два пішоходи у коридорі постійно ухиляються в один бік. Вирішення: рандомна затримка перед retry, exponential backoff.'),
  
    h2('5. Starvation (голодування)'),
    para('Потік ніколи або дуже рідко отримує CPU/ресурс через постійне витіснення потоками з вищим пріоритетом або нечесним замком. Вирішення: ReentrantLock(true) (fair lock) гарантує FIFO-чергу, уникати різниці пріоритетів там де є конкуренція.'),
  
    h2('6. Race Condition'),
    para('Некоректний результат через несподіваний порядок виконання паралельних операцій. Класичний приклад — check-then-act без атомарності:'),
    code('// НЕ атомарно:'),
    code('if (map.containsKey(key)) { return map.get(key); } // race!'),
    code('// Атомарно:'),
    code('return map.computeIfAbsent(key, k -> computeValue(k));'),
  
    h2('7. False Sharing (хибне спільне використання кешу)'),
    para('Два потоки змінюють різні змінні, що лежать в одній кеш-лінії (зазвичай 64 байти). Процесор змушений синхронізувати кеш-лінію між ядрами — суттєве сповільнення. Вирішення: @Contended аннотація (Java 8+) або ручне вирівнювання до 64 байт (padding).'),
    pageBreak(),
  
    // ── Q22 ──────────────────────────────────────────────────────────────────
    h1('Питання 22. Локери та управління потоками'),
    divider(),
  
    h2('1. java.util.concurrent.locks — огляд'),
    para('Пакет locks надає більш гнучкі та функціональні альтернативи synchronized: можливість спроби захоплення з таймаутом, переривання під час очікування, справедлива черга, кілька умовних змінних на одному замку.'),
  
    h2('2. ReentrantLock'),
    para('Реентерабельний ексклюзивний замок. Ідентичний семантиці synchronized, але з розширеними можливостями:'),
    code('ReentrantLock lock = new ReentrantLock();'),
    code('lock.lock();'),
    code('try { criticalSection(); }'),
    code('finally { lock.unlock(); } // ЗАВЖДИ у finally!'),
    bullet('tryLock() — спробувати без очікування, повертає boolean'),
    bullet('tryLock(timeout, unit) — спробувати з таймаутом'),
    bullet('lockInterruptibly() — захопити, але з можливістю переривання'),
    bullet('getHoldCount() — кількість вкладених захоплень (реентерабельність)'),
    bullet('isLocked(), isHeldByCurrentThread() — стан замку'),
  
    h2('3. ReentrantReadWriteLock'),
    para('Дозволяє кілька одночасних читачів або одного записувача. Підвищує продуктивність для структур з частим читанням і рідким записом:'),
    code('ReentrantReadWriteLock rwLock = new ReentrantReadWriteLock();'),
    code('Lock readLock  = rwLock.readLock();'),
    code('Lock writeLock = rwLock.writeLock();'),
    code(''),
    code('// Читання (кілька паралельно):'),
    code('readLock.lock(); try { return data.get(key); } finally { readLock.unlock(); }'),
    code(''),
    code('// Запис (ексклюзивно):'),
    code('writeLock.lock(); try { data.put(key, value); } finally { writeLock.unlock(); }'),
  
    h2('4. StampedLock (Java 8+)'),
    para('Оптимістичний замок для читання: tryOptimisticRead() повертає stamp. Якщо запис стався між читанням — validate(stamp) повертає false і треба перечитати з read lock. Ефективніший за ReadWriteLock при низькій конкуренції:'),
    code('long stamp = sl.tryOptimisticRead();'),
    code('int value = readData();'),
    code('if (!sl.validate(stamp)) {'),
    code('    stamp = sl.readLock();'),
    code('    try { value = readData(); } finally { sl.unlockRead(stamp); }'),
    code('}'),
  
    h2('5. Condition — умовні змінні'),
    para('lock.newCondition() повертає Condition, прив\'язаний до замку. await() = аналог wait(), signal() = notify(), signalAll() = notifyAll(). Перевага: кілька умовних змінних на одному замку (відокремлення "not full" і "not empty" у Producer-Consumer).'),
  
    h2('6. AbstractQueuedSynchronizer (AQS)'),
    para('Базовий клас для більшості java.util.concurrent примітивів: ReentrantLock, Semaphore, CountDownLatch, CyclicBarrier. Реалізує FIFO-чергу очікувальних потоків через CAS на int state. Для власних синхронізаторів — наслідувати AQS та реалізувати tryAcquire/tryRelease.'),
    pageBreak(),
  
    // ── Q24 ──────────────────────────────────────────────────────────────────
    h1('Питання 24. Інтерфейс Executor та бібліотечні класи, що його реалізують'),
    divider(),
  
    h2('1. Ієрархія інтерфейсів'),
    para('Executor (базовий) → ExecutorService (управління + Future) → ScheduledExecutorService (розклад). Реалізації: ThreadPoolExecutor, ForkJoinPool, ScheduledThreadPoolExecutor.'),
  
    h2('2. Executor — мінімальний інтерфейс'),
    code('public interface Executor {'),
    code('    void execute(Runnable command);'),
    code('}'),
    para('Декуплює подання задачі від механізму її виконання. execute() може виконати задачу в новому потоці, пулі, або навіть у поточному (CallerRunsPolicy).'),
  
    h2('3. ExecutorService — повноцінний сервіс'),
    bullet('submit(Runnable) → Future<?> — подати і отримати Future для моніторингу'),
    bullet('submit(Callable<T>) → Future<T> — з результатом'),
    bullet('invokeAll(Collection<Callable<T>>) → List<Future<T>> — всі задачі'),
    bullet('invokeAny(Collection<Callable<T>>) → T — перша завершена'),
    bullet('shutdown() — зупинити після завершення поточних задач'),
    bullet('shutdownNow() → List<Runnable> — interrupt + повернути очікуючі'),
    bullet('awaitTermination(time, unit) — чекати завершення після shutdown'),
  
    h2('4. Executors — фабричний клас'),
    code('// Фіксований пул (для CPU-bound задач):'),
    code('ExecutorService fixed = Executors.newFixedThreadPool(Runtime.getRuntime().availableProcessors());'),
    code(''),
    code('// Кешований пул (для I/O-bound задач, короткі бурсти):'),
    code('ExecutorService cached = Executors.newCachedThreadPool();'),
    code(''),
    code('// Один потік (гарантована послідовність):'),
    code('ExecutorService single = Executors.newSingleThreadExecutor();'),
    code(''),
    code('// Розклад:'),
    code('ScheduledExecutorService sched = Executors.newScheduledThreadPool(2);'),
    code('sched.scheduleAtFixedRate(task, 0, 1, TimeUnit.SECONDS);'),
    code('sched.scheduleWithFixedDelay(task, 0, 500, TimeUnit.MILLISECONDS);'),
  
    h2('5. ThreadPoolExecutor — деталі'),
    para('Параметри конструктора: corePoolSize (мінімум "живих" потоків), maximumPoolSize (максимум), keepAliveTime (TTL для idle-потоків понад core), workQueue (черга задач), threadFactory, rejectionHandler.'),
    para('Rejection policies: AbortPolicy (default, кидає), CallerRunsPolicy (виконує у відправнику — backpressure), DiscardPolicy (ігнорує), DiscardOldestPolicy (видаляє найстарішу).'),
  
    h2('6. CompletableFuture (Java 8+)'),
    para('Асинхронні обчислення з ланцюжками:'),
    code('CompletableFuture'),
    code('  .supplyAsync(() -> fetchData(), executor)'),
    code('  .thenApply(data -> process(data))'),
    code('  .thenCombine(other, (a, b) -> merge(a, b))'),
    code('  .exceptionally(ex -> fallback())'),
    code('  .thenAccept(result -> send(result));'),
    pageBreak(),
  
    // ── Q26 ──────────────────────────────────────────────────────────────────
    h1('Питання 26. Методи моделювання паралельних обчислень'),
    divider(),
  
    h2('1. Навіщо моделювати'),
    para('Перед написанням коду паралельного алгоритму потрібно зрозуміти: чи є паралелізм, де залежності між задачами, яка ефективність очікується. Моделі дозволяють це зробити формально, без виконання коду.'),
  
    h2('2. Граф DAG (Directed Acyclic Graph)'),
    para('Задачі — вузли, залежності — дуги. Дуга A→B означає: B починається лише після A. Незалежні вузли (немає шляху між ними) можуть виконуватись паралельно. Критичний шлях = нижня межа T_∞. Паралелізм = T₁/T_∞.'),
    para('Застосування: аналіз паралелізму ForkJoin-алгоритмів, планування задач у розподілених системах, аналіз компіляторних оптимізацій.'),
  
    h2('3. Мережі Петрі'),
    para('Формальна модель: місця (places) — стан/ресурси, переходи (transitions) — дії, токени (tokens) — ресурси/потоки, дуги — умови. Перехід спрацьовує (fires), коли у всіх вхідних місцях є токени. Потужна модель для аналізу deadlock, liveness, безпеки (safety) паралельних систем.'),
  
    h2('4. CSP (Communicating Sequential Processes)'),
    para('Алгебра Хоара (1978). Процеси взаємодіють через синхронні канали. Формальна верифікація можлива (model checker FDR4). Основа для Go goroutines, Erlang, Clojure core.async. Дозволяє доводити відсутність deadlock.'),
  
    h2('5. Модель акторів (Actor Model)'),
    para('Actori — незалежні сутності з приватним станом, що спілкуються лише через асинхронні повідомлення. Немає спільної пам\'яті → немає race conditions за визначенням. Реалізації: Erlang/OTP, Akka (JVM), .NET Orleans.'),
  
    h2('6. Модель BSP (Bulk Synchronous Parallel)'),
    para('Обчислення розбиваються на супер-кроки: фаза локальних обчислень → фаза комунікацій → бар\'єрна синхронізація. Простий аналіз вартості. Основа для Apache Spark (RDD transformations + shuffle).'),
  
    h2('7. Аналітичне моделювання (Amdahl, Gustafson)'),
    para('Математичні формули для оцінки прискорення та ефективності без запуску програми. Використовуються на стадії проектування для обґрунтування доцільності розпаралелення.'),
    pageBreak(),
  
    // ── Q28 ──────────────────────────────────────────────────────────────────
    h1('Питання 28. Моделювання паралельних програм мережею Петрі'),
    divider(),
  
    h2('1. Базові поняття мережі Петрі'),
    para('Мережа Петрі (Petri Net, 1962) — математична модель паралельних та розподілених систем. Компоненти:'),
    bullet('Місця (Places) — кола; представляють стани, умови або буфери ресурсів'),
    bullet('Переходи (Transitions) — прямокутники; представляють події або дії'),
    bullet("Дуги (Arcs) — спрямовані; з'єднують місця з переходами та навпаки (не безпосередньо місця між собою)"),
    bullet('Токени (Tokens) — крапки у місцях; представляють наявність ресурсу або потоку'),
    para('Розмітка (Marking) — розподіл токенів по місцях у момент часу t.'),
  
    h2('2. Правила спрацювання'),
    para('Перехід T enabled (може спрацювати), якщо у КОЖНОМУ вхідному місці є хоча б один токен (або стільки, скільки вказано на дузі — вагові дуги). При спрацюванні: видаляє по одному токену з вхідних місць, додає по одному токену до вихідних. Кілька enabled переходів можуть спрацювати паралельно (concurrent firing).'),
  
    h2('3. Моделювання паралелізму'),
    para('AND-розгалуження (fork): перехід T має кілька вихідних місць. Після спрацювання T токени з\'являються у всіх вихідних місцях одночасно — моделюємо паралельний запуск потоків.'),
    para('AND-злиття (join): перехід T має кілька вхідних місць. Спрацьовує лише коли всі вхідні місця мають токени — моделюємо join() (очікування всіх потоків).'),
  
    h2('4. Моделювання Producer-Consumer мережею Петрі'),
    para('Місця: "buffer free" (N токенів = ємність), "buffer occupied" (0 токенів), "producer ready", "consumer ready". Переходи: "produce" (з "producer ready" + "buffer free" → "buffer occupied" + "producer ready"), "consume" (з "consumer ready" + "buffer occupied" → "buffer free" + "consumer ready"). Взаємне виключення при доступі до буфера: місце "mutex" з 1 токеном — захоплюється і звільняється навколо produce/consume.'),
  
    h2('5. Аналіз властивостей'),
    bullet('Safety (безпека): жодне місце не матиме > k токенів → немає переповнення буфера'),
    bullet('Liveness (живість): кожен перехід може знову спрацювати → немає deadlock'),
    bullet('Reachability: чи можна досягти певної розмітки → аналіз можливих станів'),
    bullet('Boundedness: граф досяжності скінченний → система не виходить з-під контролю'),
  
    h2('6. Інструменти'),
    para('PIPE (Platform Independent Petri Net Editor), GreatSPN, Tina — для малювання, симуляції та аналізу мереж Петрі. JFLAP — навчальний інструмент. Формальна верифікація через state space exploration.'),
    pageBreak(),
  
    // ── Q30 ──────────────────────────────────────────────────────────────────
    h1('Питання 30. Алгоритми паралельного підсумовування та оцінка їх ефективності'),
    divider(),
  
    h2('1. Задача'),
    para('Обчислити S = a[0] + a[1] + ... + a[N-1] паралельно. Послідовно: O(N) операцій, T₁ = N-1 додавань.'),
  
    h2('2. Наївний рівномірний поділ (Linear Decomposition)'),
    para('Масив ділиться на p рівних частин. Кожен потік t підсумовує свою частину [t*N/p, (t+1)*N/p). Після завершення — головний потік підсумовує p часткових результатів:'),
    code('// Кожен потік:'),
    code('double localSum = 0;'),
    code('for (int i = start; i < end; i++) localSum += a[i];'),
    code('partialSums[threadId] = localSum;'),
    code('// Після join всіх:'),
    code('double total = 0;'),
    code('for (double ps : partialSums) total += ps;'),
    para('Tₚ ≈ N/p + p. Оптимальний p = √N дає Tₚ = 2√N. Прискорення S(p) = N / (N/p + p). При p=√N: S_max ≈ √N/2.'),
  
    h2('3. Дерево підсумовування (Tree Reduction)'),
    para('Бінарне дерево глибиною log₂(N). На кожному рівні — N/2^l паралельних додавань. Після log₂(N) кроків — результат. T_∞ = log₂(N). Максимальний паралелізм = N/log₂(N).'),
    code('// ForkJoin реалізація — класичний приклад:'),
    code('class SumTask extends RecursiveTask<Long> {'),
    code('    protected Long compute() {'),
    code('        if (hi - lo <= THRESHOLD) return sequentialSum();'),
    code('        int mid = (lo + hi) >>> 1;'),
    code('        SumTask left = new SumTask(a, lo, mid);'),
    code('        left.fork();'),
    code('        long right = new SumTask(a, mid, hi).compute();'),
    code('        return right + left.join();'),
    code('    }'),
    code('}'),
    para('Tₚ ≈ N/p + log₂(p). Прискорення ≈ p (майже лінійне для великих N).'),
  
    h2('4. Scan (Prefix Sum) — паралельний'),
    para('Обчислити масив B[i] = a[0]+a[1]+...+a[i] (running sum). Паралельний алгоритм: 2 фази — up-sweep (побудова дерева) і down-sweep (розповсюдження). T_∞ = O(log N), T₁ = O(N), ефективність E = 1. Реалізовано у GPU-шейдерах та Java parallelPrefix.'),
    code('// Java 8+:'),
    code('Arrays.parallelPrefix(array, (a, b) -> a + b);'),
  
    h2('5. Оцінка ефективності'),
    bullet('Наївний поділ: S(p) ≈ p при N >> p; E(p) ≈ 1'),
    bullet('Tree reduction: S(p) = N/(N/p + log p); при p = N/log N → S_max = log N'),
    bullet('Накладні витрати: fork/join overhead при малих THRESHOLD'),
    bullet('Кеш-ефект: послідовне читання масиву дуже кеш-ефективне; обидва алгоритми мають хорошу cache locality'),
    pageBreak(),
  
    // ── Q32 ──────────────────────────────────────────────────────────────────
    h1('Питання 32. Експериментальне дослідження ефективності паралельних обчислень'),
    divider(),
  
    h2('1. Методологія'),
    para('Коректне вимірювання продуктивності Java вимагає: JVM warmup (кілька "прогрівних" запусків для JIT-компіляції), кілька повторень (10–30 ітерацій), обчислення медіани та стандартного відхилення, ізоляції від зовнішніх факторів (фоновий GC, system load).'),
  
    h2('2. JMH (Java Microbenchmark Harness)'),
    para('Стандартний інструмент для Java benchmarking. Автоматично обробляє warmup, dead code elimination, measurement iterations:'),
    code('@BenchmarkMode(Mode.AverageTime)'),
    code('@OutputTimeUnit(TimeUnit.MILLISECONDS)'),
    code('@Warmup(iterations = 5)'),
    code('@Measurement(iterations = 10)'),
    code('@Fork(2)'),
    code('public class MatrixMultiplyBenchmark {'),
    code('    @Param({"1", "2", "4", "8"})'),
    code('    int threads;'),
    code('    @Benchmark'),
    code('    public double[][] multiply(Blackhole bh) {'),
    code('        return parallelMultiply(A, B, threads);'),
    code('    }'),
    code('}'),
  
    h2('3. Що вимірювати'),
    bullet('Elapsed time (wall-clock) — реальний час від старту до фінішу'),
    bullet('CPU time — сума часу CPU усіх потоків (CPU time / elapsed = ефективне використання ядер)'),
    bullet('Throughput — задач/секунду'),
    bullet('Latency — час однієї задачі'),
  
    h2('4. Побудова speedup-кривої'),
    para('Варіюємо p від 1 до max_cores. Вимірюємо T₁ та Tₚ. Будуємо S(p) = T₁/Tₚ та E(p) = S(p)/p. Порівнюємо з теоретичною кривою Амдала:'),
    code('double T1 = measure(1);'),
    code('for (int p : new int[]{1, 2, 4, 8, 16}) {'),
    code('    double Tp = measure(p);'),
    code('    System.out.printf("p=%d  S=%.2f  E=%.2f%%%n", p, T1/Tp, 100*T1/(p*Tp));'),
    code('}'),
  
    h2('5. Типові результати для множення матриць'),
    para('При N=1000, p=8 (8-ядерна машина): очікуване S(8) ≈ 7–7.5, E(8) ≈ 87–93%. Зниження ефективності: накладні витрати на Thread.join(), false sharing у матриці C, нерівний поділ рядків при N%p ≠ 0.'),
  
    h2('6. MPI_Wtime для вимірювання в MPI'),
    code('double t_start = MPI_Wtime();'),
    code('// ... обчислення ...'),
    code('double t_local = MPI_Wtime() - t_start;'),
    code('double t_max;'),
    code('MPI_Reduce(&t_local, &t_max, 1, MPI_DOUBLE, MPI_MAX, 0, comm);'),
    code('if (rank == 0) printf("Time: %.3f sec\\n", t_max);'),
    pageBreak(),
  
    // ── Q34 ──────────────────────────────────────────────────────────────────
    h1('Питання 34. Проектування паралельних програм'),
    divider(),
  
    h2('1. Методологія PCAM (Foster, 1995)'),
    para('Систематичний процес розробки паралельних програм з 4 кроків:'),
    bullet('Partitioning (розбиття) — розбити задачу на якомога дрібніші незалежні частини (задачі і дані)'),
    bullet('Communication (комунікація) — визначити потоки даних між частинами; мінімізувати обсяг і частоту комунікацій'),
    bullet('Agglomeration (агломерація) — об\'єднати дрібні задачі у великі для зменшення накладних витрат і покращення кеш-локальності'),
    bullet('Mapping (відображення) — призначити агломеровані задачі конкретним процесорам/потокам; балансувати навантаження'),
  
    h2('2. Крок 1: Partitioning'),
    para('Domain decomposition: розбивати дані (матриці — рядками, стовпцями, блоками). Functional decomposition: розбивати різні функції (pipeline). Правило: число задач >> число процесорів — це дає гнучкість планувальнику.'),
  
    h2('3. Крок 2: Communication'),
    para('Local communication: кожна задача спілкується з невеликою кількістю сусідів (сіткові задачі, алгоритм Кеннона — тільки циклічні зсуви). Global communication: всі спілкуються з усіма (MPI_Allreduce, MPI_Allgather) — дорого, уникати де можливо.'),
  
    h2('4. Крок 3: Agglomeration'),
    para('Збільшуємо зерно паралелізму: замість N задач (по одному рядку) → N/p задач (по групі рядків). Ціль: T_comp >> T_comm (compute-to-communication ratio >> 1). Для матричного множення: блок N/√p рядків замість 1 рядка.'),
  
    h2('5. Крок 4: Mapping'),
    para('Static mapping: задачі рівномірно розподіляються заздалегідь (підходить для регулярних задач — множення матриць). Dynamic mapping (work stealing): ForkJoinPool, ExecutorService з чергою. Cyclic mapping: задача t → процесор (t mod p) — для рівномірного розподілу при нерівномірному навантаженні.'),
  
    h2('6. Практичний приклад: Spring Boot матричне множення'),
    para('Partitioning: рядки матриці A → задачі. Communication: A — read-only shared; C — незалежні сегменти. Agglomeration: кожна задача = N/numCores рядків. Mapping: ThreadPoolExecutor з N/numCores задачами на numCores потоків.'),
    pageBreak(),
  
    // ── Q36 ──────────────────────────────────────────────────────────────────
    h1('Питання 36. Моделі пам\'яті паралельних обчислень'),
    divider(),
  
    h2('1. Що таке модель пам\'яті'),
    para('Модель пам\'яті (Memory Model) — специфікація того, як і коли записи одного потоку/процесу стають видимими іншим. Без чіткої моделі — компілятор і CPU можуть переставляти операції для оптимізації, що ламає паралельний код.'),
  
    h2('2. Sequential Consistency (SC)'),
    para('Найсуворіша модель: результат паралельного виконання еквівалентний якомусь послідовному (серіалізованому) виконанню, де кожен потік зберігає відносний порядок своїх операцій. Проста для розуміння, але дорога у реалізації (заборонено переупорядкування). Апаратно: SPARC TSO наближається, x86 має total store order.'),
  
    h2('3. Release Consistency'),
    para('Гарантує: всі записи до release (unlock) видимі після acquire (lock). Між acquire і release — немає гарантій. Ефективніший за SC. Основа для Java synchronized: вхід у synchronized = acquire fence, вихід = release fence.'),
  
    h2('4. Java Memory Model (JMM)'),
    para('JMM (JSR-133, Java 5+) визначає happens-before (HB) відношення. Якщо дія A HB дія B — то B бачить всі ефекти A. Правила HB:'),
    bullet('Порядок програми: у тому ж потоці A перед B → A HB B'),
    bullet('Monitor unlock HB наступний lock на тому ж моніторі'),
    bullet('Write volatile HB наступний read тієї ж volatile змінної'),
    bullet('Thread.start() HB будь-яка дія у запущеному потоці'),
    bullet('Будь-яка дія у потоці HB Thread.join() цього потоку'),
    bullet('Транзитивність: A HB B і B HB C → A HB C'),
  
    h2('5. Volatile у Java'),
    para('volatile гарантує: видимість (write visible to all threads), заборону переупорядкування через volatile-змінну (full memory fence), але НЕ атомарність складених операцій (i++ не атомарний навіть для volatile int). Використання: прапори завершення, published references (singleton double-check locking).'),
    code('// Коректний double-check locking:'),
    code('private volatile Singleton instance;'),
    code('public Singleton get() {'),
    code('    if (instance == null)'),
    code('        synchronized (this) {'),
    code('            if (instance == null)'),
    code('                instance = new Singleton();'),
    code('        }'),
    code('    return instance;'),
    code('}'),
  
    h2('6. MPI Memory Model'),
    para('У MPI немає спільної пам\'яті між процесами. Видимість досягається лише через явний обмін повідомленнями. MPI_Send/MPI_Recv гарантують: записи до MPI_Send видимі після MPI_Recv на стороні отримувача (message passing = synchronization point).'),
    pageBreak(),
  
    // ── Q38 ──────────────────────────────────────────────────────────────────
    h1('Питання 38. Стандарт Message Passing Interface (MPI)'),
    divider(),
  
    h2('1. Призначення та історія'),
    para('MPI — стандартний інтерфейс передачі повідомлень для паралельних програм на системах з розподіленою пам\'яттю. Розроблений MPI Forum у 1994 (MPI-1), MPI-2 (1997) додав динамічні процеси, MPI-3 (2012) — нові колективні операції. Реалізації: OpenMPI, MPICH, Intel MPI, MS-MPI.'),
  
    h2('2. Модель виконання SPMD'),
    para('Single Program, Multiple Data: всі процеси запускають один і той самий виконавчий файл, але гілкуються за rank. Запускається mpirun/mpiexec:'),
    code('mpirun -np 8 ./my_program   # 8 процесів'),
  
    h2('3. Основні функції ініціалізації'),
    code('MPI_Init(&argc, &argv);              // ініціалізувати MPI'),
    code('MPI_Comm_rank(MPI_COMM_WORLD, &rank); // мій номер'),
    code('MPI_Comm_size(MPI_COMM_WORLD, &size); // загальна кількість'),
    code('MPI_Finalize();                       // завершити MPI'),
  
    h2('4. Point-to-Point комунікації'),
    para('Блокуючі: MPI_Send, MPI_Recv. Неблокуючі: MPI_Isend, MPI_Irecv + MPI_Wait/MPI_Test. Комбіновані: MPI_Sendrecv (одночасна відправка і прийом — запобігає deadlock). Режими: стандартний, буферизований (MPI_Bsend), синхронний (MPI_Ssend), ready (MPI_Rsend).'),
  
    h2('5. Колективні операції'),
    bullet('MPI_Bcast — broadcast від root до всіх'),
    bullet('MPI_Scatter — розподіл рівних частин від root до всіх'),
    bullet('MPI_Gather — збір від всіх до root'),
    bullet('MPI_Allgather — Gather + broadcast результату всім'),
    bullet('MPI_Reduce — редукція (SUM, MAX, MIN, PROD...) до root'),
    bullet('MPI_Allreduce — Reduce + broadcast до всіх'),
    bullet('MPI_Barrier — синхронізація всіх процесів'),
    bullet('MPI_Alltoall — кожен надсилає кожному'),
    bullet('MPI_Scan — parallel prefix (prefix sum) по процесах'),
  
    h2('6. Типи даних MPI'),
    para('Базові: MPI_INT, MPI_DOUBLE, MPI_FLOAT, MPI_CHAR, MPI_LONG, MPI_BYTE. Похідні: MPI_Type_contiguous (масив однотипних), MPI_Type_vector (з кроком — для стовпців матриці), MPI_Type_struct (гетерогенна структура). Похідні типи дозволяють передавати нерегулярні структури даних одним викликом.'),
  
    h2('7. Комунікатори'),
    para('MPI_COMM_WORLD — всі процеси. MPI_Comm_split() — розбити на підгрупи. MPI_Cart_create() — топологія решітки з логічними координатами. Комунікатори дозволяють ізолювати колективні операції для підгруп (рядок/стовпець решітки в алгоритмі Фокса/Кеннона).'),
    pageBreak(),
  
    // ── Q40 ──────────────────────────────────────────────────────────────────
    h1('Питання 40. Колективні методи MPI та їх застосування для паралельних програм'),
    divider(),
  
    h2('1. Загальні властивості колективних операцій'),
    para('Колективні операції охоплюють ВСІХ процесів комунікатора — не можна викликати лише частиною. Є неявним бар\'єром (для більшості реалізацій). Внутрішньо оптимізовані (логарифмічні алгоритми, InfiniBand hardware collectives).'),
  
    h2('2. MPI_Bcast — широкомовна розсилка'),
    para('Процес root розсилає свій буфер всім іншим. Використання: розсилати матрицю B усім процесам перед стрічковим множенням, розсилати параметри алгоритму:'),
    code('if (rank == 0) initMatrix(B, N, N);'),
    code('MPI_Bcast(B, N*N, MPI_DOUBLE, 0, MPI_COMM_WORLD);'),
  
    h2('3. MPI_Scatter / MPI_Gather'),
    para('Scatter: root надсилає різним процесам різні частини масиву (рівні блоки). Gather: зворотна операція — збирає від всіх до root:'),
    code('// Scatter рядків матриці A:'),
    code('MPI_Scatter(A, rowsPerProc*M, MPI_DOUBLE,'),
    code('            localA, rowsPerProc*M, MPI_DOUBLE, 0, comm);'),
    code('// ... обчислення ...'),
    code('MPI_Gather(localC, rowsPerProc*K, MPI_DOUBLE,'),
    code('           C, rowsPerProc*K, MPI_DOUBLE, 0, comm);'),
  
    h2('4. MPI_Reduce / MPI_Allreduce'),
    para('Reduce: агрегує значення від всіх процесів до root (сума, max, min тощо). Allreduce: результат у ВСІХ процесах (без додаткового Bcast). Використання: обчислення глобальної суми, max помилки, перевірка збіжності:'),
    code('double localNorm = computeLocalNorm();'),
    code('double globalNorm;'),
    code('MPI_Allreduce(&localNorm, &globalNorm, 1, MPI_DOUBLE, MPI_SUM, comm);'),
    code('globalNorm = sqrt(globalNorm); // у всіх процесах'),
  
    h2('5. MPI_Allgather та MPI_Alltoall'),
    para('Allgather: кожен процес отримує всі елементи від всіх. Використання: збір розподіленого вектору для ітеративного алгоритму. Alltoall: транспозиція матриці розподілу — кожен надсилає різну частину кожному. Дорога операція O(p²) даних; уникати при великих p.'),
  
    h2('6. MPI_Scan — префіксна сума'),
    para('Процес i отримує редукцію елементів від процесів 0..i. Аналог Arrays.parallelPrefix але для розподілених процесів:'),
    code('int localCount = countLocalItems();'),
    code('int prefixSum;'),
    code('MPI_Scan(&localCount, &prefixSum, 1, MPI_INT, MPI_SUM, comm);'),
    code('int globalOffset = prefixSum - localCount; // стартовий індекс для мого сегмента'),
  
    h2('7. Оптимізація колективних операцій'),
    para('Алгоритм Bcast: binomial tree O(log p) кроків. Reduce: теж binomial tree. Allreduce: ring-allreduce (використовується у ML-тренуванні) — рівномірне навантаження O(2*(p-1)*N/p). Векторні варіанти: MPI_Scatterv, MPI_Gatherv для нерівних частин.'),
    pageBreak(),
  
    // ── Q42 ──────────────────────────────────────────────────────────────────
    h1('Питання 42. Розробка ефективних паралельних алгоритмів в OpenMPI'),
    divider(),
  
    h2('1. Мінімізація комунікаційних витрат'),
    para('Головний принцип: T_comp >> T_comm. Для досягнення: збільшувати зерно (розмір блоків), замінювати Allreduce → Reduce + Bcast де можливо, використовувати неблокуючі операції для перекриття комунікацій і обчислень.'),
    code('// Перекриття compute + communicate:'),
    code('MPI_Irecv(nextBlock, ..., &recvReq);  // почати отримання'),
    code('compute(currentBlock);                 // обчислювати поточний'),
    code('MPI_Wait(&recvReq, MPI_STATUS_IGNORE); // дочекатись наступного'),
  
    h2('2. Балансування навантаження'),
    para('Статичний: рівномірний поділ даних (N/p рядків кожному). Якщо N не ділиться на p — останній процес отримує залишок; або ліпше: перші N%p процесів беруть по N/p+1 рядків. Динамічний: master-worker з чергою задач — master розподіляє, workers запитують нові задачі після завершення.'),
  
    h2('3. Оптимізація Scatter/Gather для матриць'),
    para('MPI_Scatter відправляє рядки матриці: кожен процес отримує rowsPerProc суміжних рядків — добра locality. Для передачі стовпців потрібен MPI_Type_vector (stride-тип) або транспозиція перед Scatter.'),
    code('// Стовпці матриці через MPI_Type_vector:'),
    code('MPI_Type_vector(N, 1, N, MPI_DOUBLE, &colType);'),
    code('MPI_Type_commit(&colType);'),
    code('MPI_Scatter(A, 1, colType, localCol, N, MPI_DOUBLE, 0, comm);'),
  
    h2('4. Топологічна оптимізація'),
    para('MPI_Cart_create() + MPI_Cart_shift() дозволяє ефективно виконувати зсуви по решітці (алгоритм Кеннона) — MPI бере до уваги фізичну топологію і може оптимізувати маршрутизацію. MPI_Dims_create() автоматично вибирає оптимальний розмір решітки для заданого p.'),
  
    h2('5. Persistent communication'),
    para('Для повторюваних обмінів у циклі — persistent requests ефективніші (уникають повторної ініціалізації):'),
    code('MPI_Send_init(buf, N, MPI_DOUBLE, dest, tag, comm, &req);'),
    code('for (int step = 0; step < steps; step++) {'),
    code('    updateBuffer(buf);'),
    code('    MPI_Start(&req);'),
    code('    compute();'),
    code('    MPI_Wait(&req, MPI_STATUS_IGNORE);'),
    code('}'),
    code('MPI_Request_free(&req);'),
  
    h2('6. Профілювання MPI-програм'),
    para('PMPI (Profiling MPI) — стандартний шар для перехоплення MPI-викликів. Інструменти: Score-P, TAU, Intel Trace Analyzer, Vampir (візуалізація трасувань). Метрики: T_comm/T_total, load imbalance, MPI_Barrier wait time.'),
    pageBreak(),
  
    // ── Q44 ──────────────────────────────────────────────────────────────────
    h1('Питання 44. Паралельна реалізація алгоритму МГУА (метод групового урахування аргументів)'),
    divider(),
  
    h2('1. Що таке МГУА'),
    para('МГУА (Метод Групового Урахування Аргументів, GMDH — Group Method of Data Handling) — еволюційний алгоритм машинного навчання, розроблений О.Г. Івахненком у 1968. Будує поліноміальну нейромережу шляхом автоматичного відбору та комбінування предикторів.'),
    para('Ключова ідея: на кожному шарі перебираються всі пари входів, будуються регресійні моделі (поліноми Колмогорова-Габора), відбираються найкращі за критерієм (MSE на валідаційній вибірці), виходи кращих моделей стають входами наступного шару. Ітерується до погіршення зовнішнього критерію.'),
  
    h2('2. Обчислювальна складність'),
    para('Для N вхідних змінних на першому шарі: C(N,2) = N*(N-1)/2 пар → для N=100 це 4950 моделей. Кожна модель — розв\'язок системи рівнянь методом найменших квадратів (МНК): O(m³) де m — кількість вимірювань. Загалом: O(N² * m³) — чудовий кандидат для розпаралелення.'),
  
    h2('3. Рівні паралелізму'),
    para('Рівень 1 — паралельне навчання моделей у шарі: кожна пара (i,j) утворює незалежну модель → незалежні задачі для різних потоків/процесів. Рівень 2 — паралельне оцінювання на тестовій вибірці. Рівень 3 — ensemble МГУА з різними початковими параметрами.'),
  
    h2('4. Java-реалізація (ForkJoin)'),
    code('class LayerTask extends RecursiveTask<List<Model>> {'),
    code('    protected List<Model> compute() {'),
    code('        // Генерувати всі пари входів:'),
    code('        List<Pair> pairs = generatePairs(inputs);'),
    code('        // Паралельно навчати:'),
    code('        List<RecursiveTask<Model>> tasks = pairs.stream()'),
    code('            .map(p -> new ModelFitTask(p, trainData))'),
    code('            .collect(toList());'),
    code('        invokeAll(tasks);'),
    code('        // Відібрати найкращі:'),
    code('        return tasks.stream()'),
    code('            .map(ForkJoinTask::join)'),
    code('            .sorted(Comparator.comparingDouble(Model::getValidationError))'),
    code('            .limit(selectionSize)'),
    code('            .collect(toList());'),
    code('    }'),
    code('}'),
  
    h2('5. MPI-реалізація'),
    para('Process 0 (master): генерує пари, розподіляє через MPI_Scatter, збирає результати через MPI_Gather, відбирає найкращі, розсилає виходи наступного шару через MPI_Bcast. Processes 1..p-1 (workers): отримують блок пар, навчають моделі, повертають (коефіцієнти + validation MSE).'),
  
    h2('6. Прискорення та масштабованість'),
    para('Алгоритм МГУА є embarrassingly parallel на рівні пар у шарі: незалежні задачі, мінімальна комунікація (тільки Scatter/Gather між шарами). Теоретичне прискорення ≈ p (ефективність > 90% при p << числа пар). Основний overhead: MPI_Gather великих моделей та сортування результатів на master.'),
    pageBreak(),
  
    // ── Q46 ──────────────────────────────────────────────────────────────────
    h1('Питання 46. Технологія Remote Method Invocation (RMI)'),
    divider(),
  
    h2('1. Що таке RMI'),
    para('Java RMI (Remote Method Invocation) — технологія, що дозволяє об\'єктам JVM викликати методи об\'єктів, що знаходяться в іншій JVM (на тому ж або іншому комп\'ютері). Абстракція: виклик виглядає як звичайний метод, але мережевий транспорт прихований. Введена у Java 1.1 (1997).'),
  
    h2('2. Архітектура RMI'),
    para('Компоненти: Remote Interface (інтерфейс, що наслідує java.rmi.Remote), Remote Object (реалізація на сервері, наслідує UnicastRemoteObject), Stub (заглушка на клієнті — проксі, що серіалізує аргументи і надсилає), Skeleton (на стороні сервера — десеріалізує і викликає реальний метод), RMI Registry (сервіс іменування — сервер реєструє об\'єкти, клієнт шукає).'),
  
    h2('3. Реалізація RMI-сервісу'),
    code('// 1) Інтерфейс:'),
    code('public interface MatrixService extends Remote {'),
    code('    double[][] multiply(double[][] A, double[][] B) throws RemoteException;'),
    code('}'),
    code(''),
    code('// 2) Реалізація (сервер):'),
    code('public class MatrixServiceImpl extends UnicastRemoteObject'),
    code('                               implements MatrixService {'),
    code('    public double[][] multiply(double[][] A, double[][] B) {'),
    code('        return parallelMultiply(A, B); // ForkJoin всередині'),
    code('    }'),
    code('}'),
    code(''),
    code('// 3) Реєстрація:'),
    code('MatrixService svc = new MatrixServiceImpl();'),
    code('Registry reg = LocateRegistry.createRegistry(1099);'),
    code('reg.bind("MatrixService", svc);'),
    code(''),
    code('// 4) Клієнт:'),
    code('Registry reg = LocateRegistry.getRegistry("server-host", 1099);'),
    code('MatrixService svc = (MatrixService) reg.lookup("MatrixService");'),
    code('double[][] C = svc.multiply(A, B); // прозорий виклик'),
  
    h2('4. Серіалізація'),
    para('Аргументи і повернені значення мають реалізовувати Serializable. Для double[][] — вбудована серіалізація Java. При великих матрицях — значний overhead на серіалізацію/десеріалізацію. Альтернатива для production: gRPC + Protocol Buffers (набагато швидше).'),
  
    h2('5. RMI vs REST vs gRPC'),
    bullet('RMI: Java-only, статична прив\'язка, прозорість виклику, але складний firewall traversal, залежить від Java'),
    bullet('REST/JSON: universally supported, простий, але baggy (text-based), немає типобезпеки'),
    bullet('gRPC/Protobuf: бінарний, швидкий, polyglot, streaming, type-safe через IDL'),
    bullet('Spring Boot + REST — стандарт для мікросервісів замість RMI'),
    pageBreak(),
  
    // ── Q48 ──────────────────────────────────────────────────────────────────
    h1('Питання 48. Базові складові грід-системи'),
    divider(),
  
    h2('1. Визначення грід-системи'),
    para('Грід-система (Grid System) — розподілена інфраструктура, що об\'єднує обчислювальні ресурси (CPU, пам\'ять, сховище, ПЗ) різних організацій, географічно розподілених, для вирішення масштабних наукових і бізнесових задач. Ресурси різнорідні (heterogeneous) і адміністровані різними організаціями.'),
  
    h2('2. Обчислювальні ресурси (Compute Resources)'),
    para('Кластери та суперкомп\'ютери: основна обчислювальна потужність грід. Десктопи та робочі станції: волонтерські грід (BOINC). Cloud instances: сучасні грід-системи інтегрують IaaS (AWS, Azure, Google Cloud). Основна характеристика: heterogeneity — різні OS (Linux, Windows), процесори (x86, ARM), кількість ядер.'),
  
    h2('3. Системи зберігання даних (Storage Resources)'),
    para('SAN (Storage Area Network), NAS (Network Attached Storage), паралельні файлові системи: Lustre (WLCG), GPFS/SpectrumScale (IBM HPC), HDFS (Hadoop/Spark). GridFTP — протокол передачі великих файлів між вузлами грід (до 10 Gbps, multiple TCP streams, partial file access).'),
  
    h2('4. Мережева інфраструктура'),
    para('Виробнича мережа: GÉANT (Євросоюз), Internet2 (США), URAN (Україна) — виділені наукові мережі для грід. Між вузлами кластеру: InfiniBand (56–400 Gbps), Ethernet 10/25/100 Gbps. Протоколи: TCP/IP (основа), GridFTP, HTTPS, GSI (Grid Security).'),
  
    h2('5. Middleware (проміжне ПЗ)'),
    para('Middleware — шар ПЗ між ресурсами і користувачами, що приховує гетерогенність:'),
    bullet('Globus Toolkit: GSI (аутентифікація), GRAM (управління задачами), GridFTP'),
    bullet('gLite/ARC/DIRAC: middleware для WLCG'),
    bullet('Condor/HTCondor: управління навантаженням'),
    bullet('UNICORE: безпечний доступ до HPC-ресурсів'),
  
    h2('6. Сервіси безпеки'),
    para('X.509 сертифікати і PKI: кожен користувач і ресурс має сертифікат. Proxy-сертифікати: тимчасові делеговані credentials (24–168 годин) для запуску задач без введення пароля. VOMS (Virtual Organization Membership Service): авторизація на основі ролей у VO.'),
  
    h2('7. Сервіси інформації та моніторингу'),
    para('Information Services: LDAP-базовані каталоги ресурсів (MDS, BDII). Реєструють: доступну потужність, стан черг, ПЗ на вузлах. Моніторинг: Nagios, Ganglia, Grafana — стан вузлів, черг, завдань. Resource Broker — автоматично вибирає найкращі вузли для задачі.'),
    pageBreak(),
  
    // ── Q50 ──────────────────────────────────────────────────────────────────
    h1('Питання 50. Організація і управління розподіленими грід-ресурсами'),
    divider(),
  
    h2('1. Моделі управління ресурсами'),
    para('Централізована: єдиний менеджер ресурсів розподіляє задачі (простіше, але вузьке місце). Ієрархічна: локальні менеджери вузлів + глобальний брокер (стандарт для великих грід). Децентралізована (peer-to-peer): вузли самоорганізовуються (BOINC, BitTorrent-подібні). Федеративна: кілька незалежних грід-доменів з угодами про взаємодію (EGEE, OSG).'),
  
    h2('2. Resource Broker та Job Scheduling'),
    para('Resource Broker (брокер ресурсів) — ключовий компонент: отримує JDL (Job Description Language) від користувача → запитує Information Service про доступні ресурси → вибирає оптимальний вузол (matchmaking) → надсилає задачу через GRAM. Matchmaking: порівняння вимог задачі (ОС, ПЗ, RAM, CPUs) з можливостями вузлів.'),
    code('// Приклад JDL (Job Description Language):'),
    code('Executable = "matrix_multiply";'),
    code('Arguments = "1024 8";'),
    code('Requirements = (Memory >= 4096) && (CPUs >= 8);'),
    code('OutputSandbox = {"output.dat", "stderr.log"};'),
  
    h2('3. Планувальники задач'),
    para('SLURM (Simple Linux Utility for Resource Management): найпопулярніший у HPC. Підтримує: пріоритизацію, backfill scheduling, резервування, масові задачі (array jobs), гібридне MPI+GPU. PBS/Torque: класичний академічний планувальник. LSF (IBM Spectrum LSF): комерційний, для фінансового сектору. Grid Engine (UGE): підходить для гетерогенних середовищ.'),
    code('#!/bin/bash'),
    code('#SBATCH --job-name=matrix_mul'),
    code('#SBATCH --nodes=4'),
    code('#SBATCH --ntasks-per-node=16'),
    code('#SBATCH --mem=32GB'),
    code('#SBATCH --time=02:00:00'),
    code('#SBATCH --partition=hpc'),
    code('module load openmpi/4.1'),
    code('mpirun -np 64 ./matrix_multiply 4096'),
  
    h2('4. Data Management у грід'),
    para('Репліки даних: GridFTP реплікує файли між вузлами для locality (задача виконується де є дані, а не навпаки). LFC (LCG File Catalog) та Rucio (CERN) — каталоги репліків: de-facto де знаходиться кожен файл. DPM (Disk Pool Manager) — управління дисковим простором на вузлах. Storage Element (SE) vs Computing Element (CE) — ролі вузлів у класичній грід-архітектурі.'),
  
    h2('5. Моніторинг та SLA'),
    para('GridICE, GSTAT, Dashboard — web-інтерфейси моніторингу WLCG. Метрики: availability (% uptime), reliability (% успішних задач), throughput (задач/день). SLA (Service Level Agreement) між організаціями: гарантований % ресурсів для VO. Типова угода WLCG: site зобов\'язується надати >= X CPUs і Y TB storage VO CMS або ATLAS.'),
  
    h2('6. Сучасні тенденції: Cloud + Grid'),
    para('Хмарні платформи (AWS, GCP, Azure) все більше використовуються як Grid-вузли: spot instances для HPC-пакетних задач, Kubernetes для контейнеризованих обчислень, HTCondor-on-Cloud для burst computing. EGI Federated Cloud (Євросоюз) — гібридна грід+хмарна інфраструктура для наукових досліджень. Open Science Grid (OSG, США) — подібна ініціатива з інтеграцією AWS і Google Cloud.'),
  
    h2('7. Управління Virtual Organizations'),
    para('VO Management: VOMS-Admin — web-інтерфейс для адміністратора VO. Користувачі реєструються, адміністратор затверджує членство. Ролі: VO Admin, Production Role, Software Manager. Grid Certificate: від акредитованого CA (EUGridPMA, IGTF). Термін дії proxy: 24–168 годин. voms-proxy-init генерує proxy з атрибутами VO для подачі задач.'),
  );
  
  // ─── assemble ────────────────────────────────────────────────────────────────
  
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 24 } },
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 36, bold: true, font: 'Arial', color: '1F3864' },
          paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 },
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: 'Arial', color: '2E75B6' },
          paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 1000, after: 400 },
          children: [new TextRun({
            text: 'Технологія паралельних обчислень',
            bold: true, size: 48, font: 'Arial', color: '1F3864',
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 200 },
          children: [new TextRun({
            text: 'Відповіді до екзамену — Парні питання (2–50)',
            size: 32, font: 'Arial', color: '2E75B6',
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 1200 },
          children: [new TextRun({
            text: 'Java · OpenMPI · ForkJoin · Spring Boot · Grid',
            size: 24, font: 'Arial', italics: true, color: '595959',
          })],
        }),
        new Paragraph({ children: [new PageBreak()] }),
        ...sections,
      ],
    }],
  });
  
  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync('./docs/ev5.docx', buf);
  }).catch(err => console.error('❌ Помилка:', err));