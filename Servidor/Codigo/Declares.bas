Attribute VB_Name = "Declaraciones"

Option Explicit

'ANTI TIRADA DE LOGIN
Public Const MAX_CONEX As Long = 999
Public Socket_LAST As Long
Private Type SocketINFO
    ip As String
    Conectadas As Long
    Desconectadas As Long
End Type
Public Anti_DDOS(1 To MAX_CONEX) As SocketINFO
'ANTI TIRADA DE LOGIN

Public ContadorSiin As Integer
Public terminodeat As Boolean
Public CantidadEXP As Integer
Public CantidadOro As Integer

Public TiempoPremium As Integer
Public TiempoEmperador As Byte
Public MinutosTorneo As Byte
Public MinCasti As Byte
Public MinutosDeath As Byte

Public Type Tques
    NPCs As Byte
    Users As Byte
    iNPCs As Integer
    Oro As Long
    Exp As Long
    Namex As String
    Recompense As String
    Canje As Long
    Premium As Integer
End Type

Public MaxQuest As Byte

Public Type tQuestUser
    IndexNPC As Integer
    Index As Byte  'global
    NPCs As Byte    'npcs faltan
    Users As Byte   'users :P
End Type

Public Quest() As Tques


Public Const MapaX As Byte = 51
Public Const MapaY As Byte = 32
Public Const NpcRey As Integer = 257
Public HayPareja As Boolean
Public Enum trigger
    ZONAPELEA = 7
End Enum
Public Enum Trigger7
    TRIGGER7_PERMITE = 1
    TRIGGER7_PROHIBE = 2
    TRIGGER7_AUSENTE = 3
End Enum


Public Lac_Camina As Long
Public Lac_Pociones As Long
Public Lac_Pegar As Long
Public Lac_Lanzar As Long
Public Lac_Usar As Long
Public Lac_Tirar As Long
Public Type TLac
    LCaminar As New Cls_InterGTC
    LPociones As New Cls_InterGTC
    LPegar As New Cls_InterGTC
    LUsar As New Cls_InterGTC
    LTirar As New Cls_InterGTC
    LLanzar As New Cls_InterGTC
End Type
Public MinOC As Byte
Public DesaFiante(1 To 2) As Integer
Public DeFenZas As Integer
Public GlobalAct As Byte
Public PTorneo As Integer
Public ModoQuest As Boolean
Public Recargando As Boolean
Public MapDatFile As String
Public AutoTorneo As Boolean
Public PosicionTorneo As Integer

'Torneo 2 vs 2
Public MegaTorneo As Boolean
Public iPosicionTorneo As Integer
'Torneo 2 vs 2
Public Mensajes(1 To 2, 1 To 23) As String
Public Armaduras(1 To 2, 1 To 3, 1 To 6, 1 To 2) As Integer

'Matute
Type tEncuestas
    activa As Byte
    Tiempo As Integer
    votosSI As Integer
    votosNP As Integer
End Type
Public encuestas As tEncuestas
'/////////////////////////////////////

Public Const CIUDADANO = 1
Public Const TRABAJADOR = 2
Public Const EXPERTO_MINERALES = 3
Public Const MINERO = 4
Public Const HERRERO = 8
Public Const EXPERTO_MADERA = 13
Public Const TALADOR = 14
Public Const CARPINTERO = 18
Public Const PESCADOR = 23
Public Const SASTRE = 27
Public Const ALQUIMISTA = 31
Public Const LUCHADOR = 35
Public Const CON_MANA = 36
Public Const HECHICERO = 37
Public Const MAGO = 38
Public Const NIGROMANTE = 39
Public Const ORDEN_SAGRADA = 40
Public Const PALADIN = 41
Public Const CLERIGO = 42
Public Const NATURALISTA = 43
Public Const BARDO = 44
Public Const DRUIDA = 45
Public Const SIGILOSO = 46
Public Const ASESINO = 47
Public Const CAZADOR = 48
Public Const SIN_MANA = 49
Public Const ARQUERO = 50
Public Const GUERRERO = 51
Public Const CABALLERO = 52
Public Const BANDIDO = 53
Public Const PIRATA = 55
Public Const LADRON = 56

Public ELUs(1 To 100) As Long
Public EFrags(1 To 101) As Long
Public EnviarEstats As Integer
Public MixedKey As Long
Public CrcSubKey As String
'Castillos
Public CastilloNorte As String
Public HoraNorte As String
Public DateNorte As String
Public CastilloSur As String
Public HoraSur As String
Public DateSur As String

'Castillos

Public PrecioCaballo As Long

Public aDos As New clsAntiDoS
Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection
Public DiasSinLluvia As Integer

Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const PARTSANGRE = 14
Public Const FXXD = 43    'Carita xD
Public Const FXFE = 40    'Carita Feliz
Public Const FXTR = 41    'Carita Triste
Public Const FXHOT = 42    'Carita Cliente
Public Const FXBUA = 44    'Carita Llorando
Public Const FXEA = 45    'Carita Groso (?)
Public Const FXCON = 46    'Carita Confundido
Public Const FXWTF = 47    'Carita WTF xD

'Sistema de Cuentas
Public PJEnCuenta As String
Public PJEnCuentaB As String

Public Llovio As Boolean
Public Const iFragataFantasmal = 87
Public Type tLlamadaGM
    usuario As String * 255
    Desc As String * 255
End Type

Public Const LimiteNewbie = 12

Public Type tCabecera
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public loguear As Byte

Public Const TIEMPO_INICIOMEDITAR = 0
Public Const TIEMPO_SACARTEXTOENCIMA = 0

Public Const NingunEscudo = 2
Public Const NingunCasco = 2

Public Const MAXMASCOTASENTRENADOR = 7

Public Const FXWARP = 1
Public Const FXCURAR = 2

Public Const FXMEDITARCHICO = 4
Public Const FXMEDITARMEDIANO = 5
Public Const FXMEDITARGRANDE = 6
Public Const FXMEDITARGIGANTE = 60
Public Const FXMEDITARNOBLE = 59
Public Const FXAPU�ALAR = 57

Public Const POSINVALIDA = 3

Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"

Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"
Public Const Anillo As Integer = "929"
Public Const NombreAnillo As String = "Anillo M�gico"

Public Const uUsuarios = 1
Public Const uNPC = 2
Public Const uUsuariosYnpc = 3
Public Const uTerreno = 4

Public Const uPropiedades = 1
Public Const uEstado = 2
Public Const uInvocacion = 4
Public Const uMaterializa = 3
Public Const uMetamorfosis = 5
Public Const uRadial = 6
Public Const uArea = 7

Public Const Dragon = 6
Public Const MATADRAGONES = 1

Public Const MAX_MENSAJES_FORO = 35

Public Const MAXUSERHECHIZOS = 35

Public Const EsfuerzoTalarGeneral = 40
Public Const EsfuerzoTalarLe�ador = 20

Public Const EsfuerzoPescarPescador = 10
Public Const EsfuerzoPescarGeneral = 30

Public Const EsfuerzoExcavarMinero = 20
Public Const EsfuerzoExcavarGeneral = 50

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const Guardias = 6

Public Const MAXREP = 6000000
Public Const MAXORO = 9000000000#
Public Const MAXEXP = 99999999

Public Const MAXATRIBUTOS = 38
Public Const MINATRIBUTOS = 6

Public Const LingoteHierro = 386
Public Const LingotePlata = 387
Public Const LingoteOro = 388
Public Const Le�a = 58
Public Const Le�aElfica = 753

Public Const Raiz = 699

Public Const PLobo = 414
Public Const POsoPardo = 415
Public Const POsoPolar = 416

Public Const MAXNPCS = 5000
Public Const MAXCHARS = 5500

Public Const HACHA_LE�ADOR = 127
Public Const PIQUETE_MINERO = 187

Public Const RED_PESCA = 543
Public Const HILAR_SASTRE = 697

Public Const DAGA = 15
Public Const FOGATA_APAG = 136
Public Const FOGATA = 63
Public Const ORO_MINA = 194
Public Const PICO_EXPERTO = 753
Public Const PLATA_MINA = 193
Public Const HIERRO_MINA = 192
Public Const MARTILLO_HERRERO = 389
Public Const SERRUCHO_CARPINTERO = 198
Public Const ObjArboles = 4
'NPCTYPE
Public miuseremail As String
Public Const NPCTYPE_COMUN = 0
Public Const NPCTYPE_REVIVIR = 1
Public Const NPCTYPE_GUARDIAS = 2
Public Const NPCTYPE_ENTRENADOR = 3
Public Const NPCTYPE_BANQUERO = 4
Public Const NPCTYPE_NOBLE = 5
Public Const NPCTYPE_APOSTADOR = 7
Public Const NPCTYPE_TIENDA = 8
Public Const NPCTYPE_CIRUJANO = 9

Public Const NPCTYPE_VIEJO = 11
Public Const NPCTYPE_DEATH = 12
Public Const NPCTYPE_BOT = 16
Public Const NPCTYPE_CONSTRUCTOR = 17
Public Const NPCTYPE_NPCQUEST = 18
Public Const NPCTYPE_VIAJERO = 19

Public Const FX_TELEPORT_INDEX = 1

Public Const MIN_APU�ALAR = 10

Public Const mapa_castilloNorte = 214
Public Const mapa_castilloSur = 216
Public Const NUMSKILLS = 22
Public Const NUMATRIBUTOS = 5
Public Const NUMCLASES = 56
Public Const NUMRAZAS = 5

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4


Public Const MAXMASCOTAS = 20


Public Const vlASALTO = 100
Public Const vlASESINO = 1000
Public Const vlARQUERO = 5
Public Const vlNoble = 5
Public Const vlLadron = 25
Public Const vlProleta = 2


Public Const iCuerpoMuerto = 8
Public Const iCabezaMuerto = 500

Public Const iORO = 12

Public Const Teleport = 378
Public Const Flecha = 480

Public Const Pescado = 139
Public Const PescadoCaro = 543
Public Const Pescado2 = 544
Public Const Pescado3 = 545
Public Const Pescado4 = 546

Public Const ManzanaNewbie = 467
Public Const AguaNewbie = 468
Public Const DagaNewbie = 460
Public Const RopaNewbieHumano = 463
Public Const RopaNewbieElfo = 464
Public Const RopaNewbieElfoOscuro = 465
Public Const RopaNewbieEnano = 466
Public Const PocionRojaNewbie = 461

Public Const PocionRojaNoCae = 754
Public Const PocionAzulNoCae = 755
Public Const PocionAmarillaNoCae = 756
Public Const PocionVerdeNoCae = 757


Public Const Magia = 1
Public Const Robar = 2
Public Const Tacticas = 3
Public Const Armas = 4
Public Const Meditar = 5
Public Const Apu�alar = 6
Public Const Ocultarse = 7
Public Const Supervivencia = 8
Public Const Talar = 9
Public Const Defensa = 10
Public Const Pesca = 11
Public Const Mineria = 12
Public Const Carpinteria = 13
Public Const Herreria = 14
Public Const Liderazgo = 15
Public Const Domar = 16
Public Const Proyectiles = 17
Public Const Wresterling = 18
Public Const Navegacion = 19
Public Const Sastreria = 20
Public Const Comerciar = 21
Public Const Resis = 22
Public Const Resistencia = 22
Public Const Invitar = 23

Public Const FundirMetal = 88

Public Const PescarR = 99

Public MensajeRepeticion As String
Public TiempoRepeticion As Integer
Public IntervaloRepeticion As Integer
Public SoporteDesactivado As Boolean

Public Const XA = 40
Public Const XD = 10
Public Const Balance = 9

Public Const fuerza = 1
Public Const Agilidad = 2
Public Const Inteligencia = 3
Public Const Carisma = 4
Public Const Constitucion = 5

Public CuentaRegresiva As Integer
Public GMCuenta As Integer
Public MinutosTLimpiar As Byte
Public minVC As Byte

Public Const AdicionalHPGuerrero = 2
Public Const AdicionalSTLadron = 3

Public Const AdicionalSTLe�ador = 23
Public Const AdicionalSTPescador = 20
Public Const AdicionalSTMinero = 25


Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1


Public Const TileSizeX = 32
Public Const TileSizeY = 32


Public Const XWindow = 17
Public Const YWindow = 13

Public Const Nivel = 1
Public Const Muertos = 2

Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_TALAR = 13
Public Const SOUND_TIJERAS = 99
Public Const SOUND_PESCAR = 14
Public Const SOUND_MINERO = 15
Public Const SND_WARP = 3
Public Const SND_PUERTA = 5
Public Const SOUND_NIVEL = 6
Public Const SND_USERMUERTE = 11
Public Const SND_IMPACTO = 10
Public Const SND_IMPACTO2 = 12
Public Const SND_LE�ADOR = 13
Public Const SND_FOGATA = 14
Public Const SOUND_SACARARMA = 25
Public Const SND_ESCUDO = 37
Public Const MARTILLOHERRERO = 41
Public Const LABUROCARPINTERO = 42
Public Const SND_CREACIONCLAN = 44
Public Const SND_ACEPTADOCLAN = 43
Public Const SND_DECLAREWAR = 45


Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 25
Public Const MAX_NPCINVENTORY_SLOTS = 30
Public Const MAXPARTYUSERS = 4

Public Const OBJTYPE_ARMADURA = 0
Public Const OBJTYPE_CASCO = 1
Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_ESCUDO = 2
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3
Public Const OBJTYPE_ARBOLES = 4
Public Const OBJTYPE_GUITA = 5
Public Const OBJTYPE_PUERTAS = 6
Public Const OBJTYPE_CONTENEDORES = 7
Public Const OBJTYPE_CARTELES = 8
Public Const OBJTYPE_LLAVES = 9
Public Const OBJTYPE_FOROS = 10
Public Const OBJTYPE_POCIONES = 11
Public Const OBJTYPE_BEBIDA = 13
Public Const OBJTYPE_LE�A = 14
Public Const OBJTYPE_FOGATA = 15
Public Const OBJTYPE_HERRAMIENTAS = 18
Public Const OBJTYPE_TELEPORT = 19
Public Const OBJTYPE_YACIMIENTO = 22
Public Const OBJTYPE_PERGAMINOS = 24
Public Const OBJTYPE_INSTRUMENTOS = 26
Public Const OBJTYPE_YUNQUE = 27
Public Const OBJTYPE_FRAGUA = 28
Public Const OBJTYPE_MINERALES = 23
Public Const OBJTYPE_PIEL = 30
Public Const OBJTYPE_BARCOS = 31
Public Const OBJTYPE_FLECHAS = 32
Public Const OBJTYPE_BOTELLAVACIA = 33
Public Const OBJTYPE_BOTELLALLENA = 34
Public Const OBJTYPE_MANCHAS = 35
Public Const OBJTYPE_RAIZ = 36
Public Const OBJTYPE_ALAS = 37
Public Const OBJTYPE_WARP = 38
Public Const OBJTYPE_Manual = 40
Public Const OBJTYPE_CA�A = 138
Public Const OBJTYPE_CUALQUIERA = 1000




Public Const FONTTYPE_SERVER = "~65~190~156~1~0"
Public Const FONTTYPE_UDP = "~255~0~0~0~0"
Public Const FONTTYPE_TALK = "~255~255~255~0~0"
Public Const FONTTYPE_ARMADA = "~12~0~240~1~0"
Public Const FONTTYPE_CAOS = "~222~22~22~1~1"
Public Const FONTTYPE_FIGHT = "~255~0~0~1~0"
Public Const FONTTYPE_ATACO = "~230~230~0~1~0"
Public Const FONTTYPE_WARNING = "~32~51~223~1~1"
Public Const FONTTYPE_INFO = "~65~190~156~1~0"
Public Const FONTTYPE_VENENO = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD = "~255~255~255~1~0"
Public Const FONTTYPE_LVLINFO = "~200~200~200~0~0"
Public Const FONTTYPE_LVLUP = "~255~150~25~1~0"
Public Const FONTTYPE_FENIX = "~255~250~55~1~0"
Public Const FONTTYPE_MUERTO = "~255~155~55~1~0"
Public Const FONTTYPE_PARTY = "~50~175~175~1~0"
Public Const FONTTYPE_GLOBAL = "~128~255~255~1~1"
Public Const FONTTYPE_CONSEJO = "~98~204~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS = "~255~60~0~1~0"
Public Const FONTTYPE_INTRO = "~0~255~0~0~0"
Public Const FONTTYPE_DUELOS = "~255~0~255~0~1"
Public Const FONTTYPE_DESAFIO = "~94~163~204~1~1"

Public Const STAT_MAXELV = 100
Public Const STAT_MAXMAN = 99999
Public Const STAT_MAXHP = 99999
Public Const STAT_MAXSTA = 999
Public Const STAT_MAXDEF = 999

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public Const SND_NODEFAULT = &H2

Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Type tHechizo
    Nombre As String
    Desc As String
    PalabrasMagicas As String

    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    EffectIndex As Integer

    Resis As Byte

    Tipo As Byte
    WAV As Integer
    FXgrh As Long
    loops As Byte

    Baculo As Byte
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer

    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer

    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer

    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer

    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer

    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    Metamorfosis As Byte
    Body As Integer
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer

    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer

    RadioX As Byte
    RadioY As Byte

    Invisibilidad As Byte
    Transforma As Byte
    Paraliza As Byte
    Nivel As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Flecha As Byte
    Morph As Byte

    Invoca As Byte
    NumNPC As Integer
    Cant As Integer

    Materializa As Byte
    ItemIndex As Byte

    MinSkill As Integer
    StaRequerido As Integer
    ManaRequerido As Integer
    Especial As Byte

    Target As Byte
End Type

Public Codifico As Integer

Type LevelSkill

    LevelValue As Integer

End Type


Public Type UserOBJ
    OBJIndex As Long
    Amount As Integer
    Equipped As Byte
    ProbTirar As Byte
End Type


Type InventarioNPC
    Object(1 To MAX_NPCINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type

Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Long
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Long
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Long
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Long
    CascoEqpSlot As Byte
    AlaEqpObjIndex As Integer
    AlaEqpSlot As Byte
    MunicionEqpObjIndex As Long
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Long
    HerramientaEqpslot As Byte
    BarcoObjIndex As Long
    BarcoSlot As Byte
    NroItems As Integer
End Type


Type Position
    X As Integer
    Y As Integer
End Type

Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Type FXdata
    Nombre As String
    GrhIndex As Long
    Delay As Integer
End Type

Type Char
    Alas As Integer
    Account As String
    Aura As Integer
    CharIndex As Integer
    Head As Integer
    Body As Integer

    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer

    FX As Integer
    loops As Integer

    Heading As Byte
End Type
Public Worldsaves As Integer


Public Type ObjData
    Aura As Integer
    name As String
    NoComerciable As Integer

    NoSeCae As Boolean
    ObjType As Integer
    SubTipo As Integer
    Dosmanos As Integer
    GrhIndex As Long
    GrhSecundario As Long
    Jerarquia As Byte

    Respawn As Byte
    Alas As Integer

    MaxItems As Integer
    Conte As Inventario
    Apu�ala As Byte

    HechizoIndex As Integer

    ForoID As String

    MinHP As Integer
    MaxHP As Integer

    WMapa As Integer
    WX As Integer
    WY As Integer
    WI As Integer

    Baculo As Byte

    MineralIndex As Integer
    LingoteInex As Integer


    proyectil As Integer
    Municion As Integer

    Crucial As Byte
    Newbie As Integer


    MinSta As Integer


    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long

    MinSkill As Integer
    LingoteIndex As Integer

    MinHit As Integer
    MaxHit As Integer

    MinHam As Integer
    MinSed As Integer

    Def As Integer
    MinDef As Integer
    MaxDef As Integer
    Ropaje As Integer

    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    Gorro As Byte
    Valor As Long

    Cerrada As Integer
    Llave As Byte
    Clave As Long

    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer

    RazaEnana As Byte
    MUJER As Byte
    HOMBRE As Byte
    Envenena As Byte

    SkillCombate As Byte
    SkillTacticas As Byte
    SkillProyectiles As Byte
    SkillApu�alar As Byte

    Resistencia As Long

    Agarrable As Byte

    ArbolElfico As Byte
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Long
    MaderaElfica As Long


    Raices As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer

    SkHerreria As Integer
    SkCarpinteria As Integer
    SkResistencia As Integer
    SkDefensa As Integer

    SkPociones As Integer
    SkSastreria As Integer

    Texto As String


    ClaseProhibida(1 To NUMCLASES) As Integer
    RazaProhibida(1 To NUMRAZAS) As Integer
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    MinInt As Integer

    Real As Integer
    Caos As Integer

End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type TiendaObj
    OBJIndex As Integer
    Precio As Long
    Amount As Integer
End Type

Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public Const MAX_TIENDA_SLOTS = 10

Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type

Type UserTienda
    Object(1 To MAX_TIENDA_SLOTS) As TiendaObj
    NpcTienda As Integer
    NroItems As Byte
    Gold As Long
End Type












Type UserStats
    NivelMaximo As Byte
    GLD As Long
    PuntosDonador As Long
    Advertencias As Integer
    Banco As Long
    MET As Integer

    MaxHP As Integer
    MinHP As Integer

    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHit As Integer
    MinHit As Integer

    MaxHam As Integer
    MinHam As Integer

    MaxAGU As Integer
    MinAGU As Integer

    Def As Integer
    Exp As Double
    ELV As Long
    FragsLVL As Long
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    NPCsMuertos As Integer
    ManaEspecial As Byte

    SkillPts As Integer
    RecompensaLevel As Integer
    VecesMurioUsuario As Integer
End Type

Public Const ROPA_DE_EQUITACION_ENANO = 167
Public Const ROPA_DE_EQUITACION_PUESTA_ENANO = 168
Public Const ROPA_DE_EQUITACION_NORMAL = 166
Public Const ROPA_DE_EQUITACION_PUESTA_NORMAL = 168

Public dMap As Byte
Public xMap_1 As Byte
Public yMap_1 As Byte
Public xMap_2 As Byte
Public yMap_2 As Byte
Public dUser As String
Public dMoney As String
Public dIndex As String
Public Pareja As Pareja

Public Type Pareja
    Jugador1 As Integer
    Jugador2 As Integer
    Jugador3 As Integer
    Jugador4 As Integer
End Type

Type UserRanking

    DuelosGanados As Integer
    TorneosGanados As Integer
    DuelosParejaGanados As Integer
    MaxRondasDesafio As Integer

End Type

Type tPremiosCanjes
    ObjName As String
    ObjIndexP As Integer
    ObjRequiere As Integer
    ObjDescripcion As String
    ObjPremium As Integer
    ObjMaxAt As Byte
    ObjMinAt As Byte
    ObjMindef As Byte
    ObjMaxdef As Byte
    ObjMinAtMag As Byte
    ObjMaxAtMag As Byte
    ObjMinDefMag As Byte
    ObjMaxDefMag As Byte
End Type

Public PremiosList() As tPremiosCanjes

Type tGuerraClanes
    NombreClan1 As String
    NombreClan2 As String
    OcupadoCVC As Boolean
    CantidadDeParticipantes1 As Byte
    CantidadDeParticipantes2 As Byte
    CantidadQueParticipa As Byte
End Type
Public CVC As tGuerraClanes

Type tDueloXItems
    Nombre1 As String
    Nombre2 As String
    Ocupado As Boolean
End Type

Public DueloXI As tDueloXItems

Type UserFlags
    'seguridad
    DataSTRINGGENM As Integer
    DataJffsdfgdrt As Integer
    DataWEJDJz As Integer
    DatagfsdewS As Integer
    DataUEUSIDx As Integer
    'seguridad
    Death As Boolean
    EnDueloPorItems As Boolean
    ViejaPos As WorldPos
    EnCVC As Boolean
    SeguroCVC As Boolean
    Guerra As Boolean
    RecibioDonacion As Byte
    'GSEC_ID As String
    MuereQuest As Long
    Questeando As Long
    Stopped As Boolean
    SuPareja As Integer
    EsperaPareja As Boolean
    EnPareja As Boolean
    Congelado As Byte
    automatico As Boolean
    PerdioCuent As Integer
    EsNoble As Byte
    DueliandoContra As String
    LeMandaronDuelo As Boolean
    UltimoEnMandarDuelo As String
    UltimoEnMandarDueloXI As String
    LeMandaronDueloXI As Boolean
    Endueloo As Boolean
    EsperandoDuelos As Boolean
    EstanDueleando As Boolean
    Oponentes As Integer
    Esperando As Boolean
    Desafiando As Boolean
    ClienteValido As Integer
    ActGlobal As Boolean
    Rdead As Integer
    DueLeanDo As Boolean
    EsperandoDuelo As Boolean
    EstaDueleando As Boolean
    Oponente As Integer
    tdead As Integer
    EsConseCaos As Integer
    EsConseReal As Integer
    Espiado(1 To 10) As Integer
    DondeTiroMap As Integer
    DondeTiroX As Integer
    DondeTiroY As Integer
    TiroPortalL As Integer
    EsperandoLista As Integer
    DenunciasInsultos As Integer
    Denuncias As Integer
    PuedeDenunciar As Boolean
    Quest As Boolean
    LastSHWarning As Single
    MesaCasino As Byte
    BonusFlecha As Boolean
    Muerto As Byte
    Escondido As Byte
    Comerciando As Boolean
    Chateando As Boolean
    UserLogged As Boolean
    Meditando As Boolean
    Montado As Byte
    CaballoMontado As Single
    ModoCombate As Boolean
    Hambre As Byte
    chit As Long
    Sed As Byte
    PuedeMoverse As Byte
    Envenenado As Byte
    Encarcelado As Byte
    Silenciado As Byte    'matute
    Whispereando As Integer
    Sacrificado As Integer
    Sacrificador As Integer
    Sacrificando As Byte
    Paralizado As Byte
    Transformado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Lanzo As Byte
    Tomo As Byte
    Party As Byte
    InfoClanEstatica As Byte
    Invisible As Byte
    GolpeoInvi As Byte

    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    HechizoMetamorfosis As Integer
    Descansar As Boolean
    Hechizo As Integer

    EnviarHerreria As Byte
    EnviarCarpinteria As Byte
    EnviarSastreria As Byte

    QuienParalizo As Integer
    Navegando As Byte
    Seguro As Boolean

    DuracionEfecto As Long
    TomoPocion As Boolean

    Ofreciente As Integer

    TargetNpc As Integer
    TargetNpcTipo As Integer
    NpcInv As Integer

    Ban As Byte
    AdministrativeBan As Byte

    TargetUser As Integer

    TargetObj As Integer
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer

    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer

    TargetObjInvIndex As Integer
    TargetObjInvslot As Byte

    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer

    ConFrio As Integer
    Recupero As Integer
    EstasEnvenenado As Single

    StatsChanged As Byte
    Privilegios As Byte

    ValCoDe As Integer
    votoencuesta As Integer

    LastMatado(0 To 2) As String

    AdminInvisible As Byte

    Protegido As Byte
    Ignorar As Byte
    OldWeaponAnim As Integer
    TimesWalk As Integer
    StartWalk As Single
    LastWalk As Single

    CountSH As Single
    IntentosCodigo As Byte
    CodigoTrabajo As Integer
    TimerCodigo As Single
    Trabajando As Byte
    OldShieldAnim As Integer
    OldCascoAnim As Integer
End Type

Type UserCounters
    CreoTeleport As Boolean
    TimeTeleport As Integer
    Sincroniza As Single
    Protegido As Byte
    IdleCount As Single
    AttackCounter As Integer
    HPCounter As Single
    STACounter As Single
    CartelStamina As Single
    CartelFrio As Single
    Frio As Single
    COMCounter As Single
    AGUACounter As Single
    Veneno As Single
    BonusFlecha As Single
    Paralisis As Single
    Transformado As Single
    Ceguera As Single
    Estupidez As Single
    Invisibilidad As Single
    PiqueteC As Long
    TiempoPena As Single
    TiempoSilenc As Single    'matute
    PenaSilenc As Single    'matute
    Pena As Single
    SendMapCounter As WorldPos
    Pasos As Integer

    Saliendo As Boolean
    Salir As Long

    tInicioMeditar As Single
    ManaAcumulado As Single
    bPuedeMeditar As Boolean
    tUltimoHechizo As Long

    TimerUsar As Long
    TimerUsarClick As Long

    LastPot As Single
    LastPotC As Single
    LastHechizo As Single
    LastTrabajo As Single
    LastGolpe As Single
    LastRobo As Single
    LastFlecha As Single
    LastOculto As Single
End Type

Public Const Neutral = 0
Public Const Real = 1
Public Const Caos = 2

Type tFacciones
    Bando As Byte
    BandoOriginal As Byte
    Matados(0 To 2) As Integer
    Ataco(1 To 2) As Byte
    Jerarquia As Byte
    Quests As Integer
    Torneos As Integer
    RMSG As Byte
End Type

Type tCaballos
    NpcNum(2) As Integer
    Num As Integer
    Agi(2) As Integer
    pos(2) As WorldPos
End Type

Type tGuild
    Seguro As Boolean
    GuildName As String
    LoDesafiaron As Boolean
    Solicitudes As Long
    SolicitudesRechazadas As Long
    echadas As Long
    VecesFueGuildLeader As Long
    YaVoto As Byte
    EsGuildLeader As Byte
    FundoClan As Byte
    ClanFundado As String
    ClanesParticipo As Long
    GuildPoints As Double
    Bando As Byte
    Guildindex As String
End Type


Type User
    Quest As tQuestUser
    Lac As TLac
    PersonalPass As Integer
    HD As String
    name As String
    ID As Long
    Recompensas(1 To 3) As Byte

    modName As String
    Password As String
    codigo As String
    IndexPJ As Long

    Char As Char
    OrigChar As Char

    Desc As String
    Clase As Byte
    Raza As Byte
    Genero As Byte
    Email As String
    Hogar As Byte

    Invent As Inventario

    TrabajoPos As Position
    pos As WorldPos

    ConnIDvalida As Boolean
    ConnID As Integer
    RDBuffer As String
    ColaSalida As New Collection
    SockPuedoEnviar As Boolean
    CommandsBuffer As New CColaArray

    Tienda As UserTienda
    BancoInvent As BancoInventario
    Counters As UserCounters

    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMascotas As Integer

    Stats As UserStats
    flags As UserFlags
    Ranking As UserRanking
    NumeroPaquetesPorMiliSec As Long
    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long
    Caballos As tCaballos

    Faccion As tFacciones
    GuildInfo As tGuild
    GuildRef As cGuild

    PrevCRC As Long
    PacketNumber As Long
    RandKey As Long

    ip As String

    ComUsu As tCOmercioUsuario
    Accounted As String
    AccountedPass As String

    AntiCuelgue As Long
    PartyIndex As Integer
End Type

Type Party
    Bando As Byte
    NroMiembros As Byte
    MiembrosIndex(1 To MAXPARTYUSERS) As Integer
End Type

Type NPCStats
    NpcCuenta As Integer
    AutoCurar As Byte
    Alineacion As Byte
    MaxHP As Long
    MinHP As Long
    MaxHit As Integer
    MinHit As Integer
    Def As Integer
    ImpactRate As Integer
End Type

Type NpcCounters
    Paralisis As Long
    TiempoExistencia As Long
End Type

Type NPCFlags
    Apostador As Byte
    TiendaUser As Integer
    PocaParalisis As Byte
    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    QuienParalizo As Integer
    OldMovement As Byte
    OldHostil As Byte

    NoMagia As Byte
    AguaValida As Byte
    TierraInvalida As Byte

    UseAINow As Boolean
    Sound As Integer
    AttackedBy As Integer
    RespawnOrigPos As Byte

    Envenenado As Byte
    Paralizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte

    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    Snd4 As Integer

End Type

Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type


Type NpcPathFindingInfo
    path() As tVertice
    Target As Position
    PathLenght As Integer
    CurPos As Integer
    TargetUser As Integer
    NoPath As Boolean







End Type


Type Npc
    name As String
    Char As Char
    Desc As String

    NPCtype As Integer
    Numero As Integer
    AutoCurar As Integer

    level As Integer

    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNpc As Long
    TipoItems As Integer

    Veneno As Byte

    pos As WorldPos
    Orig As WorldPos
    SkillDomar As Integer

    Movement As Integer
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    Inflacion As Long

    InmuneParalisis As Byte
    GiveEXP As Long
    GiveGLD As Long

    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters

    Probabilidad As Integer
    MaxRecom As Integer
    MinRecom As Integer

    Invent As InventarioNPC
    CanAttack As Byte
    VeInvis As Byte

    NroExpresiones As Byte
    Expresiones() As String

    NroSpells As Byte
    Spells() As Integer

    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer

    PFINFO As NpcPathFindingInfo
    Bot As Byte
End Type







Type MapBlock
    Blocked As Byte
    Agua As Byte
    UserIndex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    trigger As Integer

End Type

Type InfoNPCs
    Numero As Integer
    Cantidad As Byte
End Type


Type MapInfo
    NumUsers As Integer
    NumBots As Integer
    TopPunto As Byte
    LeftPunto As Byte
    Music As String
    name As String

    NPCsTeoricos(1 To 20) As InfoNPCs
    NPCsReales(1 To 20) As InfoNPCs
    MapVersion As Integer
    SeCaenItems As Byte
    Pk As Boolean
    NoMagia As Boolean
    UserIndex() As Integer
    NpcIndex() As Integer

    Terreno As String
    Zona As String
    Restringir As Boolean
    Nivel As Byte
    BackUp As Byte
End Type

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type

Public Apuestas As tAPuestas

Public ServerOnline As Boolean
Public UltimaVersion As String
Public AUVersion As String
Public BackUp As Boolean
Public Restringido As Boolean

Public Chiters As String

Type Recompensa
    SubeHP As Integer
    SubeMP As Integer
    Obj(1 To 2) As Obj
End Type

Public Type Top
    Nombre As String
    Nivel As Integer
    Muertos As Integer
    Bando As String
End Type

Public ListaRazas() As String
Public SkillsNames() As String
Public AtributosNames() As String

Public ListaClases() As String
Public ListaBandos() As String
Public Resta() As Single
Public Recompensas() As Recompensa
Public AumentoHit() As Byte

Public Const HUMANO = 1
Public Const ENANO = 2
Public Const ELFO = 3
Public Const ELFO_OSCURO = 4
Public Const GNOMO = 5

Public Const HOMBRE = 0
Public Const MUJER = 1

Public Const HOGAR_ADELAIDE = 4
Public Const HOGAR_Althalos = 1
Public Const HOGAR_Lonelerd = 3
Public Const HOGAR_Hildegard = 2
Public Const HOGAR_LINDOS = 5

Public ENDL As String
Public ENDC As String

Public recordusuarios As Long


Public IniPath As String
Public CharPath As String
Public MapPath As String
Public DatPath As String


Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos

Public NumUsers As Integer
Public NumBots As Integer
Public NumNoGMs As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public MaxUsers2 As Integer
Public HideMe As Byte
Public MySql As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public PuedeCrearPersonajes As Byte


Type InfoHerre
    Index As Integer
    Recompensa As Byte
End Type


Public UserList() As User
Public UserOffline As User
Public Tops(1 To 2, 1 To 20) As Top
Public Npclist() As Npc
Public Party() As Party
Public NpcNoIniciado As Npc
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 100) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As InfoHerre
Public ArmadurasHerrero() As InfoHerre
Public CascosHerrero() As Integer
Public EscudosHerrero() As Integer
Public ObjDruida() As Integer
Public ObjSastre() As Integer
Public ObjCarpintero() As InfoHerre
Public Hush() As String
Public BanIps As New Collection
Public BanHDs As New Collection


Public Hildegard As WorldPos
Public Althalos As WorldPos
Public Lonelerd As WorldPos
Public LINDOS As WorldPos
Public ADELAIDE As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos
Public EmperadorPos As WorldPos
Public MataGuardiasEmperador As Byte


Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal CrcKey As Long, ByVal CrcString As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Sub PlayWaveAPI(file As String)

    On Error Resume Next
    Dim rc As Integer

    rc = sndPlaySound(file, SND_ASYNC)

End Sub
Public Function PonerPuntos(Numero As Long) As String
    Dim i As Integer
    Dim Cifra As String

    Cifra = str(Numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)
    For i = 0 To 4
        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
            End If
        Else
            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
            End If
            Exit For
        End If
    Next

    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)

End Function
