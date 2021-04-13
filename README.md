# Tarification-Assurance-Vie
# Exercice 4: Enoncé
Un contrat prévoit le versement de C$ dans n années si le souscripteur âgé aujourd’hui de x ans est
vivant à cette date, et la somme C0$; au moment du décès à un bénéficiaire désigné si le souscripteur décède avant
n années. On utilise un taux d’actualisation de 2:5%.
1. Tarifer le contrat en prime unique pure et calculer la provision mathématique pure à la kème année (on ne
tient pas compte des frais de gestion)
— Créez une interface VBA qui calcule la prime unique pure et les provisions mathématiques pures PMk
en fonction de C, C0, n et x
— Tarifer le contrat en primes annuelles constantes égales pendant n années (en début d’année) et calculer
la provision mathématique pure à la kème année (on ne tient pas compte des frais de gestion)
— Créez une interface VBA qui calcule la prime annuelle et les provisions mathématiques pures PMk en
fonction de C, C0, n et x
Afin de déterminer les primes commerciales, on introduit les chargements ci-après :
— Une proportion g1 pour la gestion des primes prélevés pendant la durée de paiement des primes et exprimés
en % de C0 , (ce chargement n’intervient pas dans le cas d’une prime unique) ;
— Une proportion g2 pour les autres frais de gestion du contrat prélevés pendant toute la durée du contrat, et
exprimés en % de C0
— l’ensemble des chargements d’acquisition est exprimé en proportion a de la prime commerciale.
1. Donner l’expression de la prime commerciale unique p, et calculer la provision mathématique commerciale , à la kème année
2. Créez une interface VBA qui calcule la prime unique commerciale et les provisions mathématiques commerciales PMk en fonction de C, C0, n , g1, g2, a et x
3. Donner l’expression de la prime commerciale annuelle constante pendant n années et calculer la provision
mathématique commerciale, à la kème année
4. Créez une interface VBA qui calcule la prime annuelle commerciale et les provisions mathématiques
commerciales PMk en fonction de C, C0, n , g1, g2, a et x

# Exercice 2:Enoncé
Un contrat d’assurance vie différée n année est émis à (x). La prestation de décès C $, est payable
au moment du décès si le décès survient après la période différée (après n années). Les primes annuelles P sont
payables en début d’année pendant la période différée (pendant les n premières années). Si l’assuré décède dans
les n premières années, les primes lui sont retournées à la fin de l’années de décès avec un taux d’intérêt j;
j ≤ i , où i est le taux d’intérêt utilisé pour évaluer la prime.
1. Si j ≤ i et j 6= 0
(a) Déterminer la prime annuelle pure P ?
(b) Calculer la provision mathématique pure à la kème année,
2. Si j = 0
(a) Déterminer la prime annuelle pure P ?
(b) Calculer la provision mathématique pure à la kème année
Objectif:Créez une interface VBA qui calcule la prime mensuelle pure P et les provisions mathématiques pure au
kème mois.

# Exercice 1:Enoncé
Un contrat EPARGNE RETRAITE a pour objet :
1. La constitution d’une épargne par capitalisation, moyennant le versement de cotisations périodiques fixes,
en vue de servir ultérieurement une retraite sous forme de capital ou de rente ;
2. Le paiement immédiat, au(x) bénéficiaire(s) désigné(s), de l’épargne constituée, en cas de décès ou d’invalidité absolue et définitive de l’assuré avant qu’il ait atteint l’âge prévu pour la retraite (60 ans).
3. Lorsque l’assuré atteint l’âge de la retraite, il aura la faculté de choisir entre l’une des quatre options de
liquidation suivantes : Option Capital ER. Option rente viagère Rv. Option rente viagère réversible Rnv.
Option rente certaine Rc. On suppose que le taux d’actualisation est 0:02.
Objectif: Créez une interface VBA qui calcule ER Rv, Rnv, Rc et la prime périodique Pi, en fonction des cotisations,
de la périodicité des cotisations :m, des frais sur cotisations fc, des fais sur encours fe, des frais sur arrérage
f et l’âge de l’assuré x.
